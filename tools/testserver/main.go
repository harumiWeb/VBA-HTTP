package main

import (
	"context"
	"encoding/json"
	"flag"
	"fmt"
	"net"
	"net/http"
	"os"
	"os/signal"
	"syscall"
	"time"
)

func main() {
	listenAddress := flag.String("listen", "127.0.0.1:0", "TCP address to listen on")
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

	ready := map[string]any{
		"event": "ready",
		"url":   "http://" + listener.Addr().String(),
	}
	if err := json.NewEncoder(os.Stdout).Encode(ready); err != nil {
		fatal(err)
	}

	serveErrors := make(chan error, 1)
	go func() {
		serveErrors <- server.Serve(listener)
	}()

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
}

func fatal(err error) {
	fmt.Fprintln(os.Stderr, err)
	os.Exit(1)
}
