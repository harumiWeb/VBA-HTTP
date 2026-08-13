package main

import (
	"bufio"
	"bytes"
	"compress/flate"
	"compress/gzip"
	"crypto/sha256"
	"crypto/tls"
	"crypto/x509"
	"encoding/hex"
	"encoding/json"
	"io"
	"mime/multipart"
	"net"
	"net/http"
	"net/http/httptest"
	"net/url"
	"strconv"
	"sync"
	"testing"
	"time"
)

func TestSelfSignedTLSFixtureIsLoopbackOnlyAndUntrusted(t *testing.T) {
	config, err := newSelfSignedTLSConfig()
	if err != nil {
		t.Fatal(err)
	}
	if len(config.Certificates) != 1 || len(config.Certificates[0].Certificate) != 1 {
		t.Fatalf("unexpected certificate chain: %#v", config.Certificates)
	}
	certificate, err := x509.ParseCertificate(config.Certificates[0].Certificate[0])
	if err != nil {
		t.Fatal(err)
	}
	if certificate.IsCA || certificate.Subject.CommonName != "VBA-HTTP untrusted fixture" {
		// The fixture is self-signed but intentionally used as a server
		// certificate; IsCA must remain false so clients cannot treat it as a
		// trust anchor merely because it is self-signed.
		t.Fatalf("unexpected certificate constraints or subject: is_ca=%v subject=%#v", certificate.IsCA, certificate.Subject)
	}
	if len(certificate.IPAddresses) != 2 || certificate.IPAddresses[0].String() != "127.0.0.1" {
		t.Fatalf("unexpected certificate IP SANs: %#v", certificate.IPAddresses)
	}
	if len(certificate.DNSNames) != 2 || certificate.DNSNames[0] != "localhost" {
		t.Fatalf("unexpected certificate DNS SANs: %#v", certificate.DNSNames)
	}
}

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

func TestRedirectLoopIsDeterministic(t *testing.T) {
	server := httptest.NewServer(newTestServer().routes())
	defer server.Close()

	client := &http.Client{CheckRedirect: func(_ *http.Request, _ []*http.Request) error { return http.ErrUseLastResponse }}
	response, err := client.Get(server.URL + "/redirect-loop")
	if err != nil {
		t.Fatal(err)
	}
	defer response.Body.Close()
	if response.StatusCode != http.StatusFound || response.Header.Get("Location") != "/redirect-loop" {
		t.Fatalf("loop response = %d %q", response.StatusCode, response.Header.Get("Location"))
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

func TestUnicodeAndMalformedUtf8EndpointsAreDeterministic(t *testing.T) {
	server := httptest.NewServer(newTestServer().routes())
	defer server.Close()

	unicodeResponse := get(t, server.URL+"/unicode")
	defer unicodeResponse.Body.Close()
	unicodeBody, err := io.ReadAll(unicodeResponse.Body)
	if err != nil {
		t.Fatal(err)
	}
	if unicodeResponse.StatusCode != http.StatusOK || unicodeResponse.Header.Get("Content-Type") != "text/plain; charset=utf-8" || string(unicodeBody) != unicodeFixture {
		t.Fatalf("unicode response = %d %q %q", unicodeResponse.StatusCode, unicodeResponse.Header.Get("Content-Type"), unicodeBody)
	}

	malformedResponse := get(t, server.URL+"/malformed-utf8")
	defer malformedResponse.Body.Close()
	malformedBody, err := io.ReadAll(malformedResponse.Body)
	if err != nil {
		t.Fatal(err)
	}
	if malformedResponse.StatusCode != http.StatusOK || malformedResponse.Header.Get("Content-Type") != "text/plain; charset=utf-8" || !bytes.Equal(malformedBody, []byte{0xc3, 0x28}) {
		t.Fatalf("malformed response = %d %q %v", malformedResponse.StatusCode, malformedResponse.Header.Get("Content-Type"), malformedBody)
	}
}

func TestCookieEndpointsAreDeterministic(t *testing.T) {
	server := httptest.NewServer(newTestServer().routes())
	defer server.Close()

	client := &http.Client{CheckRedirect: func(_ *http.Request, _ []*http.Request) error { return http.ErrUseLastResponse }}
	setResponse, err := client.Get(server.URL + "/cookie/set")
	if err != nil {
		t.Fatal(err)
	}
	defer setResponse.Body.Close()
	if setResponse.StatusCode != http.StatusNoContent || len(setResponse.Header.Values("Set-Cookie")) != 1 {
		t.Fatalf("set response = %d, cookies = %#v", setResponse.StatusCode, setResponse.Header.Values("Set-Cookie"))
	}

	request, err := http.NewRequest(http.MethodGet, server.URL+"/cookie/echo", nil)
	if err != nil {
		t.Fatal(err)
	}
	request.Header.Set("Cookie", "session=alpha")
	echoResponse, err := client.Do(request)
	if err != nil {
		t.Fatal(err)
	}
	defer echoResponse.Body.Close()
	if echoResponse.StatusCode != http.StatusNoContent || echoResponse.Header.Get("X-Cookie-Verified") != "1" {
		t.Fatalf("echo response = %d, verified = %q", echoResponse.StatusCode, echoResponse.Header.Get("X-Cookie-Verified"))
	}

	redirectResponse, err := client.Get(server.URL + "/cookie/redirect")
	if err != nil {
		t.Fatal(err)
	}
	defer redirectResponse.Body.Close()
	if redirectResponse.StatusCode != http.StatusFound || redirectResponse.Header.Get("Location") != "/cookie/echo" {
		t.Fatalf("redirect response = %d %q", redirectResponse.StatusCode, redirectResponse.Header.Get("Location"))
	}

	clearResponse, err := client.Get(server.URL + "/cookie/clear")
	if err != nil {
		t.Fatal(err)
	}
	defer clearResponse.Body.Close()
	if clearResponse.StatusCode != http.StatusNoContent || len(clearResponse.Header.Values("Set-Cookie")) != 1 {
		t.Fatalf("clear response = %d, cookies = %#v", clearResponse.StatusCode, clearResponse.Header.Values("Set-Cookie"))
	}
}

func TestMalformedHeadersEndpointIsRejectedByNetHttpClient(t *testing.T) {
	server := httptest.NewServer(newTestServer().routes())
	defer server.Close()

	_, err := http.Get(server.URL + "/malformed-headers")
	if err == nil {
		t.Fatal("malformed response headers were accepted")
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

	proxy := httptest.NewServer(newLoopbackProxy(targetURL, target.Listener.Addr().String(), nil, false))
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

func TestLoopbackProxyBasicChallenge(t *testing.T) {
	target := httptest.NewServer(newTestServer().routes())
	defer target.Close()
	targetURL, err := url.Parse(target.URL)
	if err != nil {
		t.Fatal(err)
	}

	proxy := httptest.NewServer(newLoopbackProxy(targetURL, target.Listener.Addr().String(), nil, true))
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
	response.Body.Close()
	if response.StatusCode != http.StatusProxyAuthRequired || response.Header.Get("Proxy-Authenticate") != `Basic realm="vba-http-proxy-challenge"` {
		t.Fatalf("missing proxy challenge: status=%d challenge=%q", response.StatusCode, response.Header.Get("Proxy-Authenticate"))
	}

	request, err := http.NewRequest(http.MethodGet, target.URL+"/headers", nil)
	if err != nil {
		t.Fatal(err)
	}
	request.Header.Set("Proxy-Authorization", proxyBasicAuthorization)
	response, err = client.Do(request)
	if err != nil {
		t.Fatal(err)
	}
	defer response.Body.Close()
	if response.StatusCode != http.StatusOK || response.Header.Get("X-Test-Proxy-Forwarded") != "1" {
		t.Fatalf("authenticated proxy response = %d, forwarded = %q", response.StatusCode, response.Header.Get("X-Test-Proxy-Forwarded"))
	}
}

func TestLoopbackProxyConnectTunnelsTLS(t *testing.T) {
	target := httptest.NewTLSServer(newTestServer().routes())
	defer target.Close()
	targetURL, err := url.Parse(target.URL)
	if err != nil {
		t.Fatal(err)
	}
	owner := newTestServer()
	proxy := httptest.NewServer(newLoopbackProxy(targetURL, target.Listener.Addr().String(), owner, false))
	defer proxy.Close()

	connection := connectThroughProxy(t, proxy.Listener.Addr().String(), targetURL.Host, "")
	defer connection.Close()
	connectResponse, err := http.ReadResponse(bufio.NewReader(connection), &http.Request{Method: http.MethodConnect})
	if err != nil {
		t.Fatal(err)
	}
	connectResponse.Body.Close()
	if connectResponse.StatusCode != http.StatusOK {
		t.Fatalf("CONNECT status = %d, want %d", connectResponse.StatusCode, http.StatusOK)
	}
	secure := tls.Client(connection, &tls.Config{InsecureSkipVerify: true, ServerName: "127.0.0.1"}) // test fixture is intentionally untrusted.
	defer secure.Close()
	if err := secure.Handshake(); err != nil {
		t.Fatal(err)
	}
	request, err := http.NewRequest(http.MethodGet, target.URL+"/headers", nil)
	if err != nil {
		t.Fatal(err)
	}
	if err := request.Write(secure); err != nil {
		t.Fatal(err)
	}
	response, err := http.ReadResponse(bufio.NewReader(secure), request)
	if err != nil {
		t.Fatal(err)
	}
	defer response.Body.Close()
	if response.StatusCode != http.StatusOK {
		t.Fatalf("tunneled response status = %d, want %d", response.StatusCode, http.StatusOK)
	}
	owner.mu.Lock()
	connects := owner.proxyConnects
	authorized := owner.proxyAuthorized
	owner.mu.Unlock()
	if connects != 1 || authorized != 1 {
		t.Fatalf("CONNECT stats = attempts=%d authorized=%d, want 1/1", connects, authorized)
	}
}

func TestLoopbackProxyConnectBasicChallenge(t *testing.T) {
	target := httptest.NewTLSServer(newTestServer().routes())
	defer target.Close()
	targetURL, err := url.Parse(target.URL)
	if err != nil {
		t.Fatal(err)
	}
	owner := newTestServer()
	proxy := httptest.NewServer(newLoopbackProxy(targetURL, target.Listener.Addr().String(), owner, true))
	defer proxy.Close()

	unauthorized := connectThroughProxy(t, proxy.Listener.Addr().String(), targetURL.Host, "")
	unauthorizedResponse, err := http.ReadResponse(bufio.NewReader(unauthorized), &http.Request{Method: http.MethodConnect})
	if err != nil {
		unauthorized.Close()
		t.Fatal(err)
	}
	unauthorized.Close()
	if unauthorizedResponse.StatusCode != http.StatusProxyAuthRequired || unauthorizedResponse.Header.Get("Proxy-Authenticate") != `Basic realm="vba-http-proxy-challenge"` {
		unauthorizedResponse.Body.Close()
		t.Fatalf("CONNECT challenge = status=%d header=%q", unauthorizedResponse.StatusCode, unauthorizedResponse.Header.Get("Proxy-Authenticate"))
	}
	unauthorizedResponse.Body.Close()

	connection := connectThroughProxy(t, proxy.Listener.Addr().String(), targetURL.Host, proxyBasicAuthorization)
	defer connection.Close()
	response, err := http.ReadResponse(bufio.NewReader(connection), &http.Request{Method: http.MethodConnect})
	if err != nil {
		t.Fatal(err)
	}
	response.Body.Close()
	if response.StatusCode != http.StatusOK {
		t.Fatalf("authenticated CONNECT status = %d, want %d", response.StatusCode, http.StatusOK)
	}
	owner.mu.Lock()
	connects := owner.proxyConnects
	authorized := owner.proxyAuthorized
	owner.mu.Unlock()
	if connects != 2 || authorized != 1 {
		t.Fatalf("authenticated CONNECT stats = attempts=%d authorized=%d, want 2/1", connects, authorized)
	}
}

func connectThroughProxy(t *testing.T, proxyAddress string, targetHost string, authorization string) net.Conn {
	t.Helper()
	connection, err := net.DialTimeout("tcp", proxyAddress, 5*time.Second)
	if err != nil {
		t.Fatal(err)
	}
	request := "CONNECT " + targetHost + " HTTP/1.1\r\nHost: " + targetHost + "\r\n"
	if authorization != "" {
		request += "Proxy-Authorization: " + authorization + "\r\n"
	}
	request += "\r\n"
	if _, err := io.WriteString(connection, request); err != nil {
		connection.Close()
		t.Fatal(err)
	}
	return connection
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
		{name: "basic-challenge", path: "/auth/challenge/basic", credential: basicAuthValue, challenge: `Basic realm="vba-http-challenge"`},
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
