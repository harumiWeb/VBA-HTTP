package main

import (
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"net/url"
	"strconv"
	"strings"
	"sync"
	"time"
)

const (
	maximumBodySize = int64(2 * 1024 * 1024 * 1024)
	streamChunkSize = 64 * 1024
)

type testServer struct {
	mu                sync.Mutex
	attempts          map[string]int
	inFlight          int
	maxInFlight       int
	shutdownOnce      sync.Once
	shutdownRequested chan struct{}
}

func newTestServer() *testServer {
	return &testServer{
		attempts:          make(map[string]int),
		shutdownRequested: make(chan struct{}),
	}
}

func (s *testServer) routes() http.Handler {
	mux := http.NewServeMux()
	mux.HandleFunc("GET /healthz", s.health)
	mux.HandleFunc("POST /__admin/reset", s.reset)
	mux.HandleFunc("POST /__admin/shutdown", s.shutdown)
	mux.HandleFunc("GET /__admin/stats", s.stats)
	mux.HandleFunc("GET /status/{code}", s.status)
	mux.HandleFunc("GET /delay/{milliseconds}", s.delay)
	mux.HandleFunc("GET /disconnect", s.disconnect)
	mux.HandleFunc("GET /bytes/{size}", s.bytes)
	mux.HandleFunc("GET /stream/{size}", s.stream)
	mux.HandleFunc("GET /sha256/{size}", s.sha256)
	mux.HandleFunc("GET /headers", s.headers)
	mux.HandleFunc("POST /echo", s.echo)
	mux.HandleFunc("PUT /echo", s.echo)
	mux.HandleFunc("PATCH /echo", s.echo)
	mux.HandleFunc("DELETE /echo", s.echo)
	mux.HandleFunc("GET /redirect/{count}", s.redirect)
	mux.HandleFunc("GET /flaky/{failCount}", s.flaky)
	mux.HandleFunc("GET /rate-limit/{count}", s.rateLimit)
	mux.HandleFunc("GET /retry-status/{code}/{count}", s.retryStatus)
	mux.HandleFunc("POST /retry-status/{code}/{count}", s.retryStatus)
	mux.HandleFunc("PATCH /retry-status/{code}/{count}", s.retryStatus)
	return mux
}

func (s *testServer) shutdown(w http.ResponseWriter, _ *http.Request) {
	w.WriteHeader(http.StatusAccepted)
	s.shutdownOnce.Do(func() { close(s.shutdownRequested) })
}

func (s *testServer) health(w http.ResponseWriter, _ *http.Request) {
	writeJSON(w, http.StatusOK, map[string]string{"status": "ok"})
}

func (s *testServer) reset(w http.ResponseWriter, _ *http.Request) {
	s.mu.Lock()
	clear(s.attempts)
	s.inFlight = 0
	s.maxInFlight = 0
	s.mu.Unlock()
	w.WriteHeader(http.StatusNoContent)
}

func (s *testServer) status(w http.ResponseWriter, r *http.Request) {
	code, ok := parseBoundedInt(w, r.PathValue("code"), 200, 599, "code")
	if !ok {
		return
	}
	w.WriteHeader(code)
}

func (s *testServer) delay(w http.ResponseWriter, r *http.Request) {
	milliseconds, ok := parseBoundedInt(w, r.PathValue("milliseconds"), 0, 300_000, "milliseconds")
	if !ok {
		return
	}
	s.beginInFlight()
	defer s.endInFlight()

	timer := time.NewTimer(time.Duration(milliseconds) * time.Millisecond)
	defer timer.Stop()
	select {
	case <-r.Context().Done():
		return
	case <-timer.C:
		writeJSON(w, http.StatusOK, map[string]int{"delayed_ms": milliseconds})
	}
}

func (s *testServer) disconnect(w http.ResponseWriter, _ *http.Request) {
	hijacker, ok := w.(http.Hijacker)
	if !ok {
		http.Error(w, "connection hijacking unavailable", http.StatusInternalServerError)
		return
	}
	connection, _, err := hijacker.Hijack()
	if err == nil {
		_ = connection.Close()
	}
}

func (s *testServer) stats(w http.ResponseWriter, _ *http.Request) {
	s.mu.Lock()
	current := s.inFlight
	maximum := s.maxInFlight
	s.mu.Unlock()
	w.Header().Set("X-Current-In-Flight", strconv.Itoa(current))
	w.Header().Set("X-Max-In-Flight", strconv.Itoa(maximum))
	w.WriteHeader(http.StatusNoContent)
}

func (s *testServer) beginInFlight() {
	s.mu.Lock()
	s.inFlight++
	if s.inFlight > s.maxInFlight {
		s.maxInFlight = s.inFlight
	}
	s.mu.Unlock()
}

func (s *testServer) endInFlight() {
	s.mu.Lock()
	s.inFlight--
	s.mu.Unlock()
}

func (s *testServer) bytes(w http.ResponseWriter, r *http.Request) {
	size, ok := parseSize(w, r.PathValue("size"))
	if !ok {
		return
	}
	w.Header().Set("Content-Type", "application/octet-stream")
	w.Header().Set("Content-Length", strconv.FormatInt(size, 10))
	writePattern(w, size, false)
}

func (s *testServer) stream(w http.ResponseWriter, r *http.Request) {
	size, ok := parseSize(w, r.PathValue("size"))
	if !ok {
		return
	}
	w.Header().Set("Content-Type", "application/octet-stream")
	writePattern(w, size, true)
}

func (s *testServer) sha256(w http.ResponseWriter, r *http.Request) {
	size, ok := parseSize(w, r.PathValue("size"))
	if !ok {
		return
	}
	writeJSON(w, http.StatusOK, map[string]any{
		"algorithm": "sha256",
		"bytes":     size,
		"digest":    patternHash(size),
	})
}

func (s *testServer) headers(w http.ResponseWriter, r *http.Request) {
	writeJSON(w, http.StatusOK, map[string]any{
		"headers": r.Header,
		"method":  r.Method,
		"path":    r.URL.Path,
		"query":   r.URL.Query(),
	})
}

func (s *testServer) echo(w http.ResponseWriter, r *http.Request) {
	r.Body = http.MaxBytesReader(w, r.Body, maximumBodySize)
	if contentType := r.Header.Get("Content-Type"); contentType != "" {
		w.Header().Set("Content-Type", contentType)
	}
	w.WriteHeader(http.StatusOK)
	_, _ = io.Copy(w, r.Body)
}

func (s *testServer) redirect(w http.ResponseWriter, r *http.Request) {
	count, ok := parseBoundedInt(w, r.PathValue("count"), 0, 100, "count")
	if !ok {
		return
	}
	if count == 0 {
		writeJSON(w, http.StatusOK, map[string]int{"remaining": 0})
		return
	}

	target := fmt.Sprintf("/redirect/%d", count-1)
	if r.URL.RawQuery != "" {
		target += "?" + r.URL.RawQuery
	}
	http.Redirect(w, r, target, http.StatusFound)
}

func (s *testServer) flaky(w http.ResponseWriter, r *http.Request) {
	failCount, ok := parseBoundedInt(w, r.PathValue("failCount"), 0, 10_000, "failCount")
	if !ok {
		return
	}
	attempt := s.nextAttempt("flaky", failCount, r.URL.Query())
	if attempt <= failCount {
		writeJSON(w, http.StatusServiceUnavailable, map[string]int{"attempt": attempt})
		return
	}
	writeJSON(w, http.StatusOK, map[string]int{"attempt": attempt})
}

func (s *testServer) rateLimit(w http.ResponseWriter, r *http.Request) {
	limitCount, ok := parseBoundedInt(w, r.PathValue("count"), 0, 10_000, "count")
	if !ok {
		return
	}
	retryAfter := r.URL.Query().Get("retry_after")
	if retryAfter == "" {
		retryAfter = "1"
	}
	attempt := s.nextAttempt("rate-limit", limitCount, r.URL.Query())
	if attempt <= limitCount {
		w.Header().Set("Retry-After", retryAfter)
		writeJSON(w, http.StatusTooManyRequests, map[string]int{"attempt": attempt})
		return
	}
	writeJSON(w, http.StatusOK, map[string]int{"attempt": attempt})
}

func (s *testServer) retryStatus(w http.ResponseWriter, r *http.Request) {
	code, ok := parseBoundedInt(w, r.PathValue("code"), 400, 599, "code")
	if !ok {
		return
	}
	failCount, ok := parseBoundedInt(w, r.PathValue("count"), 0, 10_000, "count")
	if !ok {
		return
	}
	attempt := s.nextAttempt("retry-status-"+strconv.Itoa(code), failCount, r.URL.Query())
	if attempt <= failCount {
		writeJSON(w, code, map[string]int{"attempt": attempt})
		return
	}
	writeJSON(w, http.StatusOK, map[string]int{"attempt": attempt})
}

func (s *testServer) nextAttempt(endpoint string, count int, query url.Values) int {
	key := strings.Join([]string{endpoint, strconv.Itoa(count), query.Get("id")}, ":")
	s.mu.Lock()
	defer s.mu.Unlock()
	s.attempts[key]++
	return s.attempts[key]
}

func parseSize(w http.ResponseWriter, raw string) (int64, bool) {
	size, err := strconv.ParseInt(raw, 10, 64)
	if err != nil || size < 0 || size > maximumBodySize {
		writeJSON(w, http.StatusBadRequest, map[string]string{"error": "size must be between 0 and 2147483648"})
		return 0, false
	}
	return size, true
}

func parseBoundedInt(w http.ResponseWriter, raw string, minimum, maximum int, name string) (int, bool) {
	value, err := strconv.Atoi(raw)
	if err != nil || value < minimum || value > maximum {
		writeJSON(w, http.StatusBadRequest, map[string]string{
			"error": fmt.Sprintf("%s must be between %d and %d", name, minimum, maximum),
		})
		return 0, false
	}
	return value, true
}

func writePattern(w io.Writer, size int64, flush bool) {
	chunk := make([]byte, streamChunkSize)
	for index := range chunk {
		chunk[index] = byte(index % 251)
	}

	flusher, canFlush := w.(http.Flusher)
	remaining := size
	for remaining > 0 {
		writeSize := int64(len(chunk))
		if remaining < writeSize {
			writeSize = remaining
		}
		if _, err := w.Write(chunk[:writeSize]); err != nil {
			return
		}
		remaining -= writeSize
		if flush && canFlush {
			flusher.Flush()
		}
	}
}

func patternHash(size int64) string {
	hash := sha256.New()
	writePattern(hash, size, false)
	return hex.EncodeToString(hash.Sum(nil))
}

func writeJSON(w http.ResponseWriter, status int, value any) {
	w.Header().Set("Content-Type", "application/json; charset=utf-8")
	w.WriteHeader(status)
	_ = json.NewEncoder(w).Encode(value)
}
