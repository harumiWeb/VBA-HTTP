package main

import (
	"bytes"
	"compress/flate"
	"compress/gzip"
	"context"
	"crypto/sha256"
	"crypto/subtle"
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
	maximumBodySize    = int64(2 * 1024 * 1024 * 1024)
	streamChunkSize    = 64 * 1024
	compressionFixture = "VBA-HTTP compression fixture: 0123456789\n"
	unicodeFixture     = "VBA-HTTP unicode: 日本語🙂"
	basicAuthValue     = "Basic dXNlcjpwYXNz"
	bearerAuthValue    = "Bearer vba-http-token"
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
	mux.HandleFunc("GET /unicode", s.unicode)
	mux.HandleFunc("GET /malformed-utf8", s.malformedUTF8)
	mux.HandleFunc("GET /malformed-headers", s.malformedHeaders)
	mux.HandleFunc("GET /compress/gzip", s.compressGzip)
	mux.HandleFunc("GET /compress/deflate", s.compressDeflate)
	mux.HandleFunc("GET /headers", s.headers)
	mux.HandleFunc("GET /auth/basic", s.authBasic)
	mux.HandleFunc("GET /auth/bearer", s.authBearer)
	mux.HandleFunc("GET /auth/challenge/basic", s.authChallengeBasic)
	mux.HandleFunc("GET /cookie/set", s.cookieSet)
	mux.HandleFunc("GET /cookie/echo", s.cookieEcho)
	mux.HandleFunc("GET /cookie/clear", s.cookieClear)
	mux.HandleFunc("GET /cookie/redirect", s.cookieRedirect)
	mux.HandleFunc("POST /echo", s.echo)
	mux.HandleFunc("PUT /echo", s.echo)
	mux.HandleFunc("PATCH /echo", s.echo)
	mux.HandleFunc("DELETE /echo", s.echo)
	mux.HandleFunc("POST /upload/hash", s.uploadHash)
	mux.HandleFunc("POST /upload/slow/{milliseconds}", s.uploadSlow)
	mux.HandleFunc("POST /upload/multipart", s.uploadMultipart)
	mux.HandleFunc("POST /upload/challenge", s.uploadChallenge)
	mux.HandleFunc("GET /redirect/{count}", s.redirect)
	mux.HandleFunc("GET /redirect-loop", s.redirectLoop)
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

func (s *testServer) unicode(w http.ResponseWriter, _ *http.Request) {
	w.Header().Set("Content-Type", "text/plain; charset=utf-8")
	w.WriteHeader(http.StatusOK)
	_, _ = io.WriteString(w, unicodeFixture)
}

func (s *testServer) malformedUTF8(w http.ResponseWriter, _ *http.Request) {
	w.Header().Set("Content-Type", "text/plain; charset=utf-8")
	w.WriteHeader(http.StatusOK)
	_, _ = w.Write([]byte{0xc3, 0x28})
}

func (s *testServer) malformedHeaders(w http.ResponseWriter, _ *http.Request) {
	hijacker, ok := w.(http.Hijacker)
	if !ok {
		http.Error(w, "connection hijacking unavailable", http.StatusInternalServerError)
		return
	}
	connection, _, err := hijacker.Hijack()
	if err != nil {
		return
	}
	defer connection.Close()
	_, _ = io.WriteString(connection, "HTTP/1.1 200 OK\r\nContent-Type: text/plain\r\nMalformed-Header\r\nContent-Length: 0\r\n\r\n")
}

func (s *testServer) compressGzip(w http.ResponseWriter, _ *http.Request) {
	writeCompressed(w, "gzip")
}

func (s *testServer) compressDeflate(w http.ResponseWriter, _ *http.Request) {
	writeCompressed(w, "deflate")
}

func writeCompressed(w http.ResponseWriter, encoding string) {
	var buffer bytes.Buffer
	var writer io.WriteCloser
	if encoding == "gzip" {
		writer = gzip.NewWriter(&buffer)
	} else {
		writer, _ = flate.NewWriter(&buffer, flate.DefaultCompression)
	}
	if _, err := writer.Write([]byte(compressionFixture)); err != nil {
		return
	}
	if err := writer.Close(); err != nil {
		return
	}
	w.Header().Set("Content-Type", "text/plain; charset=utf-8")
	w.Header().Set("Content-Encoding", encoding)
	w.Header().Set("Vary", "Accept-Encoding")
	w.Header().Set("Content-Length", strconv.Itoa(buffer.Len()))
	w.WriteHeader(http.StatusOK)
	_, _ = w.Write(buffer.Bytes())
}

func (s *testServer) headers(w http.ResponseWriter, r *http.Request) {
	writeJSON(w, http.StatusOK, map[string]any{
		"headers": r.Header,
		"method":  r.Method,
		"path":    r.URL.Path,
		"query":   r.URL.Query(),
	})
}

func (s *testServer) authBasic(w http.ResponseWriter, r *http.Request) {
	if constantTimeEqual(r.Header.Get("Authorization"), basicAuthValue) {
		w.Header().Set("X-Auth-Verified", "1")
		w.WriteHeader(http.StatusNoContent)
		return
	}
	w.Header().Set("WWW-Authenticate", `Basic realm="vba-http-test", charset="UTF-8"`)
	w.WriteHeader(http.StatusUnauthorized)
}

func (s *testServer) authBearer(w http.ResponseWriter, r *http.Request) {
	if constantTimeEqual(r.Header.Get("Authorization"), bearerAuthValue) {
		w.Header().Set("X-Auth-Verified", "1")
		w.WriteHeader(http.StatusNoContent)
		return
	}
	w.Header().Set("WWW-Authenticate", `Bearer realm="vba-http-test", error="invalid_token"`)
	w.WriteHeader(http.StatusUnauthorized)
}

func (s *testServer) authChallengeBasic(w http.ResponseWriter, r *http.Request) {
	if constantTimeEqual(r.Header.Get("Authorization"), basicAuthValue) {
		w.Header().Set("X-Auth-Verified", "1")
		w.WriteHeader(http.StatusNoContent)
		return
	}
	w.Header().Set("WWW-Authenticate", `Basic realm="vba-http-challenge"`)
	w.WriteHeader(http.StatusUnauthorized)
}

func (s *testServer) cookieSet(w http.ResponseWriter, _ *http.Request) {
	w.Header().Add("Set-Cookie", "session=alpha; Path=/cookie; Max-Age=3600")
	w.WriteHeader(http.StatusNoContent)
}

func (s *testServer) cookieEcho(w http.ResponseWriter, r *http.Request) {
	cookie, err := r.Cookie("session")
	if err == nil && cookie.Value == "alpha" {
		w.Header().Set("X-Cookie-Verified", "1")
		w.WriteHeader(http.StatusNoContent)
		return
	}
	w.WriteHeader(http.StatusUnauthorized)
}

func (s *testServer) cookieClear(w http.ResponseWriter, _ *http.Request) {
	w.Header().Add("Set-Cookie", "session=; Path=/cookie; Max-Age=0")
	w.WriteHeader(http.StatusNoContent)
}

func (s *testServer) cookieRedirect(w http.ResponseWriter, r *http.Request) {
	http.Redirect(w, r, "/cookie/echo", http.StatusFound)
}

func constantTimeEqual(actual, expected string) bool {
	return subtle.ConstantTimeCompare([]byte(actual), []byte(expected)) == 1
}

func (s *testServer) echo(w http.ResponseWriter, r *http.Request) {
	r.Body = http.MaxBytesReader(w, r.Body, maximumBodySize)
	if contentType := r.Header.Get("Content-Type"); contentType != "" {
		w.Header().Set("Content-Type", contentType)
	}
	w.WriteHeader(http.StatusOK)
	_, _ = io.Copy(w, r.Body)
}

func (s *testServer) uploadHash(w http.ResponseWriter, r *http.Request) {
	writeUploadHash(w, r, r.Body)
}

func (s *testServer) uploadSlow(w http.ResponseWriter, r *http.Request) {
	milliseconds, ok := parseBoundedInt(w, r.PathValue("milliseconds"), 0, 10_000, "milliseconds")
	if !ok {
		return
	}
	delayed := &delayedReader{reader: r.Body, delay: time.Duration(milliseconds) * time.Millisecond, context: r.Context()}
	writeUploadHash(w, r, delayed)
}

func (s *testServer) uploadChallenge(w http.ResponseWriter, _ *http.Request) {
	w.Header().Set("WWW-Authenticate", `Basic realm="vba-http-test"`)
	w.WriteHeader(http.StatusUnauthorized)
}

func (s *testServer) uploadMultipart(w http.ResponseWriter, r *http.Request) {
	r.Body = http.MaxBytesReader(w, r.Body, maximumBodySize)
	reader, err := r.MultipartReader()
	if err != nil {
		writeJSON(w, http.StatusBadRequest, map[string]string{"error": "invalid multipart body"})
		return
	}
	fields := make(map[string]string)
	type uploadedFile struct {
		Name        string `json:"name"`
		Filename    string `json:"filename"`
		ContentType string `json:"content_type"`
		Bytes       int64  `json:"bytes"`
		Digest      string `json:"sha256"`
	}
	files := make([]uploadedFile, 0)
	for {
		part, nextErr := reader.NextPart()
		if nextErr == io.EOF {
			break
		}
		if nextErr != nil {
			writeJSON(w, http.StatusBadRequest, map[string]string{"error": "invalid multipart part"})
			return
		}
		if part.FileName() == "" {
			value, readErr := io.ReadAll(io.LimitReader(part, 1<<20))
			part.Close()
			if readErr != nil {
				writeJSON(w, http.StatusBadRequest, map[string]string{"error": "could not read multipart field"})
				return
			}
			fields[part.FormName()] = string(value)
			continue
		}
		hash := sha256.New()
		bytesRead, copyErr := io.Copy(hash, part)
		part.Close()
		if copyErr != nil {
			writeJSON(w, http.StatusBadRequest, map[string]string{"error": "could not read multipart file"})
			return
		}
		files = append(files, uploadedFile{
			Name:        part.FormName(),
			Filename:    part.FileName(),
			ContentType: part.Header.Get("Content-Type"),
			Bytes:       bytesRead,
			Digest:      hex.EncodeToString(hash.Sum(nil)),
		})
	}
	if len(files) == 1 {
		w.Header().Set("X-Multipart-File-Digest", files[0].Digest)
		w.Header().Set("X-Multipart-File-Bytes", strconv.FormatInt(files[0].Bytes, 10))
		w.Header().Set("X-Multipart-Filename", files[0].Filename)
	}
	if value, exists := fields["title"]; exists {
		w.Header().Set("X-Multipart-Field-Title-UTF8", hex.EncodeToString([]byte(value)))
	}
	writeJSON(w, http.StatusOK, map[string]any{"fields": fields, "files": files})
}

func writeUploadHash(w http.ResponseWriter, r *http.Request, body io.Reader) {
	hash := sha256.New()
	bytesRead, err := io.Copy(hash, io.LimitReader(body, maximumBodySize+1))
	if err != nil {
		return
	}
	if bytesRead > maximumBodySize {
		writeJSON(w, http.StatusRequestEntityTooLarge, map[string]string{"error": "body too large"})
		return
	}
	w.Header().Set("X-Upload-Digest", hex.EncodeToString(hash.Sum(nil)))
	w.Header().Set("X-Upload-Bytes", strconv.FormatInt(bytesRead, 10))
	writeJSON(w, http.StatusOK, map[string]any{
		"algorithm":      "sha256",
		"bytes":          bytesRead,
		"digest":         hex.EncodeToString(hash.Sum(nil)),
		"content_length": r.ContentLength,
	})
}

type delayedReader struct {
	reader  io.Reader
	delay   time.Duration
	context context.Context
}

func (d *delayedReader) Read(buffer []byte) (int, error) {
	select {
	case <-d.context.Done():
		return 0, d.context.Err()
	case <-time.After(d.delay):
	}
	return d.reader.Read(buffer)
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

func (s *testServer) redirectLoop(w http.ResponseWriter, r *http.Request) {
	http.Redirect(w, r, "/redirect-loop", http.StatusFound)
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
