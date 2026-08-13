package main

import (
	"io"
	"net"
	"net/http"
	"net/url"
	"time"
)

const proxyBasicAuthorization = "Basic cHJveHktdXNlcjpwcm94eS1wYXNz"

// loopbackProxy is a deliberately small HTTP forward proxy used by integration
// tests. It only forwards to the test server's own loopback listener, so a
// malformed test cannot accidentally reach the external network.
type loopbackProxy struct {
	target           *url.URL
	targetAddress    string
	owner            *testServer
	transport        *http.Transport
	requireBasicAuth bool
}

func newLoopbackProxy(target *url.URL, targetAddress string, owner *testServer, requireBasicAuth bool) *loopbackProxy {
	return &loopbackProxy{
		target:        target,
		targetAddress: targetAddress,
		owner:         owner,
		transport: &http.Transport{
			Proxy: nil,
		},
		requireBasicAuth: requireBasicAuth,
	}
}

func (p *loopbackProxy) ServeHTTP(writer http.ResponseWriter, request *http.Request) {
	if request.Method == http.MethodConnect {
		p.serveConnect(writer, request)
		return
	}
	if p.requireBasicAuth && request.Header.Get("Proxy-Authorization") != proxyBasicAuthorization {
		writer.Header().Set("Proxy-Authenticate", `Basic realm="vba-http-proxy-challenge"`)
		writer.WriteHeader(http.StatusProxyAuthRequired)
		return
	}

	forwardURL := *request.URL
	if !forwardURL.IsAbs() {
		forwardURL.Scheme = p.target.Scheme
		forwardURL.Host = request.Host
	}
	if forwardURL.Host != p.target.Host {
		http.Error(writer, "test proxy refuses a non-loopback target", http.StatusBadGateway)
		return
	}
	forwardURL.Scheme = p.target.Scheme
	forwardURL.Host = p.target.Host

	forward := request.Clone(request.Context())
	forward.URL = &forwardURL
	forward.RequestURI = ""
	forward.Header.Del("Proxy-Connection")
	forward.Header.Del("Proxy-Authorization")
	forward.Header.Set("X-Test-Proxy-Forwarded", "1")
	forward.Host = p.target.Host

	response, err := p.transport.RoundTrip(forward)
	if err != nil {
		http.Error(writer, "test proxy forwarding failed", http.StatusBadGateway)
		return
	}
	defer response.Body.Close()
	for name, values := range response.Header {
		for _, value := range values {
			writer.Header().Add(name, value)
		}
	}
	writer.Header().Set("X-Test-Proxy-Forwarded", "1")
	writer.WriteHeader(response.StatusCode)
	_, _ = io.Copy(writer, response.Body)
}

func (p *loopbackProxy) serveConnect(writer http.ResponseWriter, request *http.Request) {
	authorized := !p.requireBasicAuth || request.Header.Get("Proxy-Authorization") == proxyBasicAuthorization
	if p.owner != nil {
		p.owner.recordProxyConnect(authorized)
	}
	if !authorized {
		writer.Header().Set("Proxy-Authenticate", `Basic realm="vba-http-proxy-challenge"`)
		writer.WriteHeader(http.StatusProxyAuthRequired)
		return
	}

	requestedHost := request.Host
	if requestedHost == "" && request.URL != nil {
		requestedHost = request.URL.Host
	}
	if requestedHost != p.target.Host {
		http.Error(writer, "test proxy refuses a non-loopback CONNECT target", http.StatusBadGateway)
		return
	}
	if p.targetAddress == "" {
		http.Error(writer, "test proxy has no CONNECT target", http.StatusBadGateway)
		return
	}

	targetConnection, err := net.DialTimeout("tcp", p.targetAddress, 5*time.Second)
	if err != nil {
		http.Error(writer, "test proxy CONNECT target failed", http.StatusBadGateway)
		return
	}

	hijacker, ok := writer.(http.Hijacker)
	if !ok {
		_ = targetConnection.Close()
		http.Error(writer, "connection hijacking unavailable", http.StatusInternalServerError)
		return
	}
	proxyConnection, _, err := hijacker.Hijack()
	if err != nil {
		_ = targetConnection.Close()
		return
	}

	defer proxyConnection.Close()
	defer targetConnection.Close()
	if _, err = io.WriteString(proxyConnection, "HTTP/1.1 200 Connection Established\r\n\r\n"); err != nil {
		return
	}

	completed := make(chan struct{}, 2)
	go proxyCopy(targetConnection, proxyConnection, completed)
	go proxyCopy(proxyConnection, targetConnection, completed)
	<-completed
}

func proxyCopy(destination net.Conn, source net.Conn, completed chan<- struct{}) {
	_, _ = io.Copy(destination, source)
	completed <- struct{}{}
}
