package main

import (
	"context"
	"crypto/tls"
	"encoding/json"
	"flag"
	"fmt"
	"net"
	"net/http"
	"net/url"
	"os"
	"os/signal"
	"syscall"
	"time"
)

func main() {
	listenAddress := flag.String("listen", "127.0.0.1:0", "TCP address to listen on")
	tlsListenAddress := flag.String("tls-listen", "", "Optional TCP address for an HTTPS listener with an untrusted self-signed certificate")
	proxyListenAddress := flag.String("proxy-listen", "", "Optional loopback address for a deterministic HTTP forward proxy")
	proxyAuthListenAddress := flag.String("proxy-auth-listen", "", "Optional loopback address for a deterministic Basic-authenticated HTTP forward proxy")
	proxyTLSListenAddress := flag.String("proxy-tls-listen", "", "Optional loopback address for a deterministic HTTPS CONNECT proxy")
	proxyTLSAuthListenAddress := flag.String("proxy-tls-auth-listen", "", "Optional loopback address for a Basic-authenticated HTTPS CONNECT proxy")
	flag.Parse()

	listener, err := net.Listen("tcp", *listenAddress)
	if err != nil {
		fatal(err)
	}
	tcpAddress, ok := listener.Addr().(*net.TCPAddr)
	if !ok || !tcpAddress.IP.IsLoopback() {
		_ = listener.Close()
		fatal(fmt.Errorf("test server must listen on a loopback address"))
	}

	testServer := newTestServer()
	server := &http.Server{
		Handler:           testServer.routes(),
		ReadHeaderTimeout: 10 * time.Second,
	}
	var tlsListener net.Listener
	var tlsServer *http.Server
	var httpsURL string
	tlsTargetAddress := ""
	if *tlsListenAddress != "" {
		tlsConfig, tlsErr := newSelfSignedTLSConfig()
		if tlsErr != nil {
			_ = listener.Close()
			fatal(tlsErr)
		}
		tlsListener, tlsErr = tls.Listen("tcp", *tlsListenAddress, tlsConfig)
		if tlsErr != nil {
			_ = listener.Close()
			fatal(tlsErr)
		}
		tlsTCPAddress, tlsOK := tlsListener.Addr().(*net.TCPAddr)
		if !tlsOK || !tlsTCPAddress.IP.IsLoopback() {
			_ = listener.Close()
			_ = tlsListener.Close()
			fatal(fmt.Errorf("test TLS server must listen on a loopback address"))
		}
		tlsServer = &http.Server{
			Handler:           testServer.routes(),
			ReadHeaderTimeout: 10 * time.Second,
		}
		httpsURL = "https://" + tlsListener.Addr().String()
		tlsTargetAddress = tlsListener.Addr().String()
	}
	var proxyListener net.Listener
	var proxyServer *http.Server
	var proxyURL string
	var proxyAuthListener net.Listener
	var proxyAuthServer *http.Server
	var proxyAuthURL string
	var proxyTLSListener net.Listener
	var proxyTLSServer *http.Server
	var proxyTLSURL string
	var proxyTLSTargetURL string
	var proxyTLSAuthListener net.Listener
	var proxyTLSAuthServer *http.Server
	var proxyTLSAuthURL string
	readyURL := "http://" + listener.Addr().String()
	proxyTargetURL := ""
	var proxyTarget *url.URL
	if *proxyListenAddress != "" || *proxyAuthListenAddress != "" {
		proxyTarget, err = url.Parse(fmt.Sprintf("http://vba-http.localhost:%d", tcpAddress.Port))
		if err != nil {
			_ = listener.Close()
			fatal(err)
		}
		proxyTargetURL = proxyTarget.String()
	}
	var proxyTLSTarget *url.URL
	if *proxyTLSListenAddress != "" || *proxyTLSAuthListenAddress != "" {
		if tlsListener == nil {
			_ = listener.Close()
			fatal(fmt.Errorf("HTTPS CONNECT proxy requires -tls-listen"))
		}
		proxyTLSTarget, err = url.Parse(fmt.Sprintf("https://vba-http.localhost:%d", tlsListener.Addr().(*net.TCPAddr).Port))
		if err != nil {
			_ = listener.Close()
			_ = tlsListener.Close()
			fatal(err)
		}
		proxyTLSTargetURL = proxyTLSTarget.String()
	}
	if *proxyListenAddress != "" {
		proxyListener, err = net.Listen("tcp", *proxyListenAddress)
		if err != nil {
			_ = listener.Close()
			fatal(err)
		}
		proxyTCPAddress, proxyOK := proxyListener.Addr().(*net.TCPAddr)
		if !proxyOK || !proxyTCPAddress.IP.IsLoopback() {
			_ = listener.Close()
			_ = proxyListener.Close()
			fatal(fmt.Errorf("test proxy must listen on a loopback address"))
		}
		// WinHTTP deliberately bypasses proxies for literal loopback hosts.  A
		// loopback-only DNS alias keeps the fixture local while exercising the
		// named-proxy path instead of silently connecting directly.
		proxyServer = &http.Server{Handler: newLoopbackProxy(proxyTarget, tcpAddress.String(), testServer, false), ReadHeaderTimeout: 10 * time.Second}
		proxyURL = "http://" + proxyListener.Addr().String()
	}
	if *proxyAuthListenAddress != "" {
		proxyAuthListener, err = net.Listen("tcp", *proxyAuthListenAddress)
		if err != nil {
			_ = listener.Close()
			if proxyListener != nil {
				_ = proxyListener.Close()
			}
			fatal(err)
		}
		proxyAuthTCPAddress, proxyAuthOK := proxyAuthListener.Addr().(*net.TCPAddr)
		if !proxyAuthOK || !proxyAuthTCPAddress.IP.IsLoopback() {
			_ = listener.Close()
			_ = proxyAuthListener.Close()
			if proxyListener != nil {
				_ = proxyListener.Close()
			}
			fatal(fmt.Errorf("test authenticated proxy must listen on a loopback address"))
		}
		proxyAuthServer = &http.Server{Handler: newLoopbackProxy(proxyTarget, tcpAddress.String(), testServer, true), ReadHeaderTimeout: 10 * time.Second}
		proxyAuthURL = "http://" + proxyAuthListener.Addr().String()
	}
	if *proxyTLSListenAddress != "" {
		proxyTLSListener, err = net.Listen("tcp", *proxyTLSListenAddress)
		if err != nil {
			_ = listener.Close()
			if tlsListener != nil {
				_ = tlsListener.Close()
			}
			fatal(err)
		}
		proxyTCPAddress, proxyOK := proxyTLSListener.Addr().(*net.TCPAddr)
		if !proxyOK || !proxyTCPAddress.IP.IsLoopback() {
			_ = listener.Close()
			_ = proxyTLSListener.Close()
			if tlsListener != nil {
				_ = tlsListener.Close()
			}
			fatal(fmt.Errorf("test HTTPS CONNECT proxy must listen on a loopback address"))
		}
		proxyTLSServer = &http.Server{Handler: newLoopbackProxy(proxyTLSTarget, tlsTargetAddress, testServer, false), ReadHeaderTimeout: 10 * time.Second}
		proxyTLSURL = "http://" + proxyTLSListener.Addr().String()
	}
	if *proxyTLSAuthListenAddress != "" {
		proxyTLSAuthListener, err = net.Listen("tcp", *proxyTLSAuthListenAddress)
		if err != nil {
			_ = listener.Close()
			if tlsListener != nil {
				_ = tlsListener.Close()
			}
			if proxyTLSListener != nil {
				_ = proxyTLSListener.Close()
			}
			fatal(err)
		}
		proxyTCPAddress, proxyOK := proxyTLSAuthListener.Addr().(*net.TCPAddr)
		if !proxyOK || !proxyTCPAddress.IP.IsLoopback() {
			_ = listener.Close()
			_ = proxyTLSAuthListener.Close()
			if tlsListener != nil {
				_ = tlsListener.Close()
			}
			if proxyTLSListener != nil {
				_ = proxyTLSListener.Close()
			}
			fatal(fmt.Errorf("test authenticated HTTPS CONNECT proxy must listen on a loopback address"))
		}
		proxyTLSAuthServer = &http.Server{Handler: newLoopbackProxy(proxyTLSTarget, tlsTargetAddress, testServer, true), ReadHeaderTimeout: 10 * time.Second}
		proxyTLSAuthURL = "http://" + proxyTLSAuthListener.Addr().String()
	}

	ready := map[string]any{
		"event": "ready",
		"url":   readyURL,
	}
	if proxyURL != "" {
		ready["proxy_url"] = proxyURL
		ready["proxy_target_url"] = proxyTargetURL
	}
	if proxyAuthURL != "" {
		ready["proxy_auth_url"] = proxyAuthURL
	}
	if proxyTLSURL != "" {
		ready["proxy_tls_url"] = proxyTLSURL
		ready["proxy_tls_target_url"] = proxyTLSTargetURL
	}
	if proxyTLSAuthURL != "" {
		ready["proxy_tls_auth_url"] = proxyTLSAuthURL
	}
	if httpsURL != "" {
		ready["https_url"] = httpsURL
	}
	if err := json.NewEncoder(os.Stdout).Encode(ready); err != nil {
		fatal(err)
	}

	serveErrors := make(chan error, 8)
	go func() {
		serveErrors <- server.Serve(listener)
	}()
	if proxyServer != nil {
		go func() {
			serveErrors <- proxyServer.Serve(proxyListener)
		}()
	}
	if proxyAuthServer != nil {
		go func() {
			serveErrors <- proxyAuthServer.Serve(proxyAuthListener)
		}()
	}
	if proxyTLSServer != nil {
		go func() {
			serveErrors <- proxyTLSServer.Serve(proxyTLSListener)
		}()
	}
	if proxyTLSAuthServer != nil {
		go func() {
			serveErrors <- proxyTLSAuthServer.Serve(proxyTLSAuthListener)
		}()
	}
	if tlsServer != nil {
		go func() {
			serveErrors <- tlsServer.Serve(tlsListener)
		}()
	}

	signals := make(chan os.Signal, 1)
	signal.Notify(signals, os.Interrupt, syscall.SIGTERM)
	defer signal.Stop(signals)

	select {
	case err := <-serveErrors:
		if err != nil && err != http.ErrServerClosed {
			fatal(err)
		}
	case <-signals:
	case <-testServer.shutdownRequested:
	}

	ctx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
	defer cancel()
	if err := server.Shutdown(ctx); err != nil {
		fatal(err)
	}
	if proxyServer != nil {
		if err := proxyServer.Shutdown(ctx); err != nil {
			fatal(err)
		}
	}
	if proxyAuthServer != nil {
		if err := proxyAuthServer.Shutdown(ctx); err != nil {
			fatal(err)
		}
	}
	if proxyTLSServer != nil {
		if err := proxyTLSServer.Shutdown(ctx); err != nil {
			fatal(err)
		}
	}
	if proxyTLSAuthServer != nil {
		if err := proxyTLSAuthServer.Shutdown(ctx); err != nil {
			fatal(err)
		}
	}
	if tlsServer != nil {
		if err := tlsServer.Shutdown(ctx); err != nil {
			fatal(err)
		}
	}
}

func fatal(err error) {
	fmt.Fprintln(os.Stderr, err)
	os.Exit(1)
}
