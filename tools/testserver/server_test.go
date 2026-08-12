package main

import (
	"bytes"
	"compress/flate"
	"compress/gzip"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"io"
	"mime/multipart"
	"net/http"
	"net/http/httptest"
	"net/url"
	"strconv"
	"sync"
	"testing"
	"time"
)

func TestStatus(t *testing.T) {
	server := httptest.NewServer(newTestServer().routes())
	defer server.Close()

	response := get(t, server.URL+"/status/418")
	defer response.Body.Close()
	if response.StatusCode != http.StatusTeapot {
		t.Fatalf("status = %d, want %d", response.StatusCode, http.StatusTeapot)
	}
}

func TestInvalidParametersReturnBadRequest(t *testing.T) {
	server := httptest.NewServer(newTestServer().routes())
	defer server.Close()

	for _, path := range []string{"/status/199", "/delay/-1", "/bytes/-1", "/redirect/101"} {
		response := get(t, server.URL+path)
		response.Body.Close()
		if response.StatusCode != http.StatusBadRequest {
			t.Fatalf("%s status = %d, want 400", path, response.StatusCode)
		}
	}
}

func TestDelay(t *testing.T) {
	server := httptest.NewServer(newTestServer().routes())
	defer server.Close()

	started := time.Now()
	response := get(t, server.URL+"/delay/20")
	defer response.Body.Close()
	if elapsed := time.Since(started); elapsed < 15*time.Millisecond {
		t.Fatalf("delay elapsed = %s, want at least 15ms", elapsed)
	}
}

func TestBytesAndStreamUseDeterministicPattern(t *testing.T) {
	server := httptest.NewServer(newTestServer().routes())
	defer server.Close()

	for _, endpoint := range []string{"bytes", "stream"} {
		t.Run(endpoint, func(t *testing.T) {
			const size = int64(200_000)
			response := get(t, server.URL+"/"+endpoint+"/"+strconv.FormatInt(size, 10))
			defer response.Body.Close()
			body, err := io.ReadAll(response.Body)
			if err != nil {
				t.Fatal(err)
			}
			if int64(len(body)) != size {
				t.Fatalf("length = %d, want %d", len(body), size)
			}
			hash := sha256.Sum256(body)
			if got := hex.EncodeToString(hash[:]); got != patternHash(size) {
				t.Fatalf("hash = %s, want %s", got, patternHash(size))
			}
		})
	}
}

func TestSHA256DescribesGeneratedPayload(t *testing.T) {
	server := httptest.NewServer(newTestServer().routes())
	defer server.Close()

	response := get(t, server.URL+"/sha256/200000")
	defer response.Body.Close()
	var payload struct {
		Algorithm string `json:"algorithm"`
		Bytes     int64  `json:"bytes"`
		Digest    string `json:"digest"`
	}
	if err := json.NewDecoder(response.Body).Decode(&payload); err != nil {
		t.Fatal(err)
	}
	if payload.Algorithm != "sha256" || payload.Bytes != 200_000 || payload.Digest != patternHash(200_000) {
		t.Fatalf("unexpected hash payload: %#v", payload)
	}
}

func TestCompressedEndpointsUseDeterministicWireEncoding(t *testing.T) {
	server := httptest.NewServer(newTestServer().routes())
	defer server.Close()

	client := &http.Client{Transport: &http.Transport{DisableCompression: true}}
	for _, test := range []struct {
		path     string
		encoding string
		open     func(io.Reader) (io.ReadCloser, error)
	}{
		{path: "/compress/gzip", encoding: "gzip", open: func(reader io.Reader) (io.ReadCloser, error) { return gzip.NewReader(reader) }},
		{path: "/compress/deflate", encoding: "deflate", open: func(reader io.Reader) (io.ReadCloser, error) { return flate.NewReader(reader), nil }},
	} {
		t.Run(test.encoding, func(t *testing.T) {
			response, err := client.Get(server.URL + test.path)
			if err != nil {
				t.Fatal(err)
			}
			defer response.Body.Close()
			wire, err := io.ReadAll(response.Body)
			if err != nil {
				t.Fatal(err)
			}
			if response.StatusCode != http.StatusOK || response.Header.Get("Content-Encoding") != test.encoding || response.Header.Get("Vary") != "Accept-Encoding" {
				t.Fatalf("status/headers = %d %q %q", response.StatusCode, response.Header.Get("Content-Encoding"), response.Header.Get("Vary"))
			}
			decoded, err := test.open(bytes.NewReader(wire))
			if err != nil {
				t.Fatal(err)
			}
			defer decoded.Close()
			body, err := io.ReadAll(decoded)
			if err != nil {
				t.Fatal(err)
			}
			if string(body) != compressionFixture {
				t.Fatalf("decoded body = %q, want %q", body, compressionFixture)
			}
		})
	}
}

func TestLoopbackProxyForwardsOnlyToTarget(t *testing.T) {
	target := httptest.NewServer(newTestServer().routes())
	defer target.Close()
	targetURL, err := url.Parse(target.URL)
	if err != nil {
		t.Fatal(err)
	}

	proxy := httptest.NewServer(newLoopbackProxy(targetURL))
	defer proxy.Close()
	proxyURL, err := url.Parse(proxy.URL)
	if err != nil {
		t.Fatal(err)
	}
	client := &http.Client{Transport: &http.Transport{Proxy: http.ProxyURL(proxyURL)}}
	response, err := client.Get(target.URL + "/headers")
	if err != nil {
		t.Fatal(err)
	}
	defer response.Body.Close()
	if response.StatusCode != http.StatusOK || response.Header.Get("X-Test-Proxy-Forwarded") != "1" {
		t.Fatalf("proxy response = %d, forwarded = %q", response.StatusCode, response.Header.Get("X-Test-Proxy-Forwarded"))
	}

	badResponse, err := client.Get("http://127.0.0.1:1/should-not-forward")
	if err != nil {
		return
	}
	defer badResponse.Body.Close()
	if badResponse.StatusCode != http.StatusBadGateway {
		t.Fatalf("non-target status = %d, want %d", badResponse.StatusCode, http.StatusBadGateway)
	}
}

func TestAdminShutdownSignalsLifecycle(t *testing.T) {
	testServer := newTestServer()
	request := httptest.NewRequest(http.MethodPost, "/__admin/shutdown", nil)
	recorder := httptest.NewRecorder()
	testServer.routes().ServeHTTP(recorder, request)
	if recorder.Code != http.StatusAccepted {
		t.Fatalf("status = %d, want %d", recorder.Code, http.StatusAccepted)
	}
	select {
	case <-testServer.shutdownRequested:
	default:
		t.Fatal("shutdown signal was not published")
	}
}

func TestHeadersAndEcho(t *testing.T) {
	server := httptest.NewServer(newTestServer().routes())
	defer server.Close()

	request, err := http.NewRequest(http.MethodGet, server.URL+"/headers?name=value", nil)
	if err != nil {
		t.Fatal(err)
	}
	request.Header.Set("X-Test-Header", "present")
	response, err := http.DefaultClient.Do(request)
	if err != nil {
		t.Fatal(err)
	}
	defer response.Body.Close()
	var payload struct {
		Headers http.Header `json:"headers"`
		Query   urlValues   `json:"query"`
	}
	if err := json.NewDecoder(response.Body).Decode(&payload); err != nil {
		t.Fatal(err)
	}
	if payload.Headers.Get("X-Test-Header") != "present" || payload.Query.Get("name") != "value" {
		t.Fatalf("unexpected headers payload: %#v", payload)
	}

	echoBody := []byte("echo payload")
	echoResponse, err := http.Post(server.URL+"/echo", "text/plain", bytes.NewReader(echoBody))
	if err != nil {
		t.Fatal(err)
	}
	defer echoResponse.Body.Close()
	actualEcho, err := io.ReadAll(echoResponse.Body)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(actualEcho, echoBody) {
		t.Fatalf("echo = %q, want %q", actualEcho, echoBody)
	}

	for _, method := range []string{http.MethodPut, http.MethodPatch, http.MethodDelete} {
		request, err := http.NewRequest(method, server.URL+"/echo", bytes.NewReader(echoBody))
		if err != nil {
			t.Fatal(err)
		}
		response, err := http.DefaultClient.Do(request)
		if err != nil {
			t.Fatal(err)
		}
		actual, readErr := io.ReadAll(response.Body)
		response.Body.Close()
		if readErr != nil {
			t.Fatal(readErr)
		}
		if response.StatusCode != http.StatusOK || !bytes.Equal(actual, echoBody) {
			t.Fatalf("%s echo = status %d body %q", method, response.StatusCode, actual)
		}
	}
}

func TestAuthEndpointsVerifyWithoutReflectingCredentials(t *testing.T) {
	server := httptest.NewServer(newTestServer().routes())
	defer server.Close()

	cases := []struct {
		name       string
		path       string
		credential string
		challenge  string
	}{
		{name: "basic", path: "/auth/basic", credential: basicAuthValue, challenge: `Basic realm="vba-http-test", charset="UTF-8"`},
		{name: "bearer", path: "/auth/bearer", credential: bearerAuthValue, challenge: `Bearer realm="vba-http-test", error="invalid_token"`},
	}
	for _, test := range cases {
		t.Run(test.name, func(t *testing.T) {
			missing := get(t, server.URL+test.path)
			missingBody, err := io.ReadAll(missing.Body)
			missing.Body.Close()
			if err != nil {
				t.Fatal(err)
			}
			if missing.StatusCode != http.StatusUnauthorized || missing.Header.Get("WWW-Authenticate") != test.challenge {
				t.Fatalf("missing auth response = %d %q", missing.StatusCode, missing.Header.Get("WWW-Authenticate"))
			}
			if bytes.Contains(missingBody, []byte(test.credential)) {
				t.Fatal("challenge body reflected the credential")
			}

			request, err := http.NewRequest(http.MethodGet, server.URL+test.path, nil)
			if err != nil {
				t.Fatal(err)
			}
			request.Header.Set("Authorization", test.credential)
			response, err := http.DefaultClient.Do(request)
			if err != nil {
				t.Fatal(err)
			}
			body, err := io.ReadAll(response.Body)
			response.Body.Close()
			if err != nil {
				t.Fatal(err)
			}
			if response.StatusCode != http.StatusNoContent || response.Header.Get("X-Auth-Verified") != "1" {
				t.Fatalf("valid auth response = %d %q", response.StatusCode, response.Header.Get("X-Auth-Verified"))
			}
			if bytes.Contains(body, []byte(test.credential)) {
				t.Fatal("success body reflected the credential")
			}
		})
	}
}

func TestUploadHashAndChallenge(t *testing.T) {
	server := httptest.NewServer(newTestServer().routes())
	defer server.Close()

	body := bytes.Repeat([]byte{0x2a}, 100_000)
	request, err := http.NewRequest(http.MethodPost, server.URL+"/upload/hash", bytes.NewReader(body))
	if err != nil {
		t.Fatal(err)
	}
	request.Header.Set("Content-Type", "application/octet-stream")
	response, err := http.DefaultClient.Do(request)
	if err != nil {
		t.Fatal(err)
	}
	defer response.Body.Close()
	if response.StatusCode != http.StatusOK {
		t.Fatalf("upload hash status = %d", response.StatusCode)
	}
	var payload struct {
		Algorithm string `json:"algorithm"`
		Bytes     int64  `json:"bytes"`
		Digest    string `json:"digest"`
	}
	if err := json.NewDecoder(response.Body).Decode(&payload); err != nil {
		t.Fatal(err)
	}
	expected := sha256.Sum256(body)
	if payload.Algorithm != "sha256" || payload.Bytes != int64(len(body)) || payload.Digest != hex.EncodeToString(expected[:]) {
		t.Fatalf("unexpected upload hash payload: %#v", payload)
	}

	challenge, err := http.Post(server.URL+"/upload/challenge", "application/octet-stream", bytes.NewReader(body))
	if err != nil {
		t.Fatal(err)
	}
	defer challenge.Body.Close()
	if challenge.StatusCode != http.StatusUnauthorized || challenge.Header.Get("WWW-Authenticate") == "" {
		t.Fatalf("challenge status/header = %d %q", challenge.StatusCode, challenge.Header.Get("WWW-Authenticate"))
	}
}

func TestUploadMultipartParsesFieldsAndFiles(t *testing.T) {
	server := httptest.NewServer(newTestServer().routes())
	defer server.Close()

	var body bytes.Buffer
	writer := multipart.NewWriter(&body)
	if err := writer.WriteField("title", "日本語"); err != nil {
		t.Fatal(err)
	}
	file, err := writer.CreateFormFile("payload", "payload.bin")
	if err != nil {
		t.Fatal(err)
	}
	fileBytes := bytes.Repeat([]byte{0x7f}, 80_000)
	if _, err := file.Write(fileBytes); err != nil {
		t.Fatal(err)
	}
	if err := writer.Close(); err != nil {
		t.Fatal(err)
	}
	response, err := http.Post(server.URL+"/upload/multipart", writer.FormDataContentType(), &body)
	if err != nil {
		t.Fatal(err)
	}
	defer response.Body.Close()
	if response.StatusCode != http.StatusOK {
		t.Fatalf("multipart status = %d", response.StatusCode)
	}
	var payload struct {
		Fields map[string]string `json:"fields"`
		Files  []struct {
			Name     string `json:"name"`
			Filename string `json:"filename"`
			Bytes    int64  `json:"bytes"`
			Digest   string `json:"sha256"`
		} `json:"files"`
	}
	if err := json.NewDecoder(response.Body).Decode(&payload); err != nil {
		t.Fatal(err)
	}
	expected := sha256.Sum256(fileBytes)
	if payload.Fields["title"] != "日本語" || len(payload.Files) != 1 || payload.Files[0].Name != "payload" || payload.Files[0].Filename != "payload.bin" || payload.Files[0].Bytes != int64(len(fileBytes)) || payload.Files[0].Digest != hex.EncodeToString(expected[:]) {
		t.Fatalf("unexpected multipart payload: %#v", payload)
	}
}

func TestDelayStatsAndDisconnect(t *testing.T) {
	server := httptest.NewServer(newTestServer().routes())
	defer server.Close()

	var group sync.WaitGroup
	for range 4 {
		group.Add(1)
		go func() {
			defer group.Done()
			response, err := http.Get(server.URL + "/delay/50")
			if err == nil {
				response.Body.Close()
			}
		}()
	}
	group.Wait()

	stats, err := http.Get(server.URL + "/__admin/stats")
	if err != nil {
		t.Fatal(err)
	}
	stats.Body.Close()
	if stats.Header.Get("X-Current-In-Flight") != "0" || stats.Header.Get("X-Max-In-Flight") != "4" {
		t.Fatalf("unexpected in-flight stats: current=%s max=%s", stats.Header.Get("X-Current-In-Flight"), stats.Header.Get("X-Max-In-Flight"))
	}

	_, err = http.Get(server.URL + "/disconnect")
	if err == nil {
		t.Fatal("disconnect endpoint should close the connection before a response")
	}
}

func TestRedirect(t *testing.T) {
	server := httptest.NewServer(newTestServer().routes())
	defer server.Close()

	response := get(t, server.URL+"/redirect/3")
	defer response.Body.Close()
	if response.StatusCode != http.StatusOK || response.Request.URL.Path != "/redirect/0" {
		t.Fatalf("status/path = %d %s", response.StatusCode, response.Request.URL.Path)
	}
}

func TestFlakyAndReset(t *testing.T) {
	testServer := newTestServer()
	server := httptest.NewServer(testServer.routes())
	defer server.Close()

	assertStatuses(t, server.URL+"/flaky/2?id=case", []int{503, 503, 200})
	resetResponse, err := http.Post(server.URL+"/__admin/reset", "application/json", nil)
	if err != nil {
		t.Fatal(err)
	}
	resetResponse.Body.Close()
	assertStatuses(t, server.URL+"/flaky/2?id=case", []int{503})
}

func TestRateLimit(t *testing.T) {
	server := httptest.NewServer(newTestServer().routes())
	defer server.Close()

	response := get(t, server.URL+"/rate-limit/1?id=case&retry_after=7")
	response.Body.Close()
	if response.StatusCode != http.StatusTooManyRequests || response.Header.Get("Retry-After") != "7" {
		t.Fatalf("status/retry-after = %d %q", response.StatusCode, response.Header.Get("Retry-After"))
	}

	response = get(t, server.URL+"/rate-limit/1?id=case&retry_after=7")
	defer response.Body.Close()
	if response.StatusCode != http.StatusOK {
		t.Fatalf("status = %d, want 200", response.StatusCode)
	}
}

func TestRetryStatusSupportsMethodAndStatusMatrix(t *testing.T) {
	server := httptest.NewServer(newTestServer().routes())
	defer server.Close()

	for _, method := range []string{http.MethodGet, http.MethodPost, http.MethodPatch} {
		target := server.URL + "/retry-status/502/1?id=" + method
		for attempt, want := range []int{http.StatusBadGateway, http.StatusOK} {
			request, err := http.NewRequest(method, target, nil)
			if err != nil {
				t.Fatal(err)
			}
			response, err := http.DefaultClient.Do(request)
			if err != nil {
				t.Fatal(err)
			}
			response.Body.Close()
			if response.StatusCode != want {
				t.Fatalf("%s attempt %d status = %d, want %d", method, attempt+1, response.StatusCode, want)
			}
		}
	}
}

type urlValues map[string][]string

func (v urlValues) Get(key string) string {
	if values := v[key]; len(values) > 0 {
		return values[0]
	}
	return ""
}

func assertStatuses(t *testing.T, target string, expected []int) {
	t.Helper()
	for _, want := range expected {
		response := get(t, target)
		response.Body.Close()
		if response.StatusCode != want {
			t.Fatalf("status = %d, want %d", response.StatusCode, want)
		}
	}
}

func get(t *testing.T, target string) *http.Response {
	t.Helper()
	response, err := http.Get(target)
	if err != nil {
		t.Fatal(err)
	}
	return response
}
