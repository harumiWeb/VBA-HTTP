package main

import (
	"io"
	"net/http"
	"net/url"
)

// loopbackProxy is a deliberately small HTTP forward proxy used by integration
// tests. It only forwards to the test server's own loopback listener, so a
// malformed test cannot accidentally reach the external network.
type loopbackProxy struct {
	target    *url.URL
	transport *http.Transport
}

func newLoopbackProxy(target *url.URL) *loopbackProxy {
	return &loopbackProxy{
		target: target,
		transport: &http.Transport{
			Proxy: nil,
		},
	}
}

func (p *loopbackProxy) ServeHTTP(writer http.ResponseWriter, request *http.Request) {
	if request.Method == http.MethodConnect {
		http.Error(writer, "CONNECT is not supported by the deterministic test proxy", http.StatusNotImplemented)
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
