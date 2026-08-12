package main

import (
	"encoding/json"
	"os"
	"path/filepath"
	"testing"
)

func TestBenchmarkSchemaAndPhaseZeroBaselinesAreValidJSON(t *testing.T) {
	for _, relativePath := range []string{
		filepath.Join("..", "..", "benchmarks", "schema", "benchmark-result.schema.json"),
		filepath.Join("..", "..", "benchmarks", "results", "raw-winhttp-baseline.json"),
		filepath.Join("..", "..", "benchmarks", "results", "vba-web-baseline.json"),
	} {
		content, err := os.ReadFile(relativePath)
		if err != nil {
			t.Fatal(err)
		}
		if !json.Valid(content) {
			t.Fatalf("%s is not strict JSON", relativePath)
		}
	}
}

func TestPhaseZeroBaselinesSatisfyRequiredContract(t *testing.T) {
	tests := []struct {
		file           string
		implementation string
		expectedCommit string
	}{
		{file: "raw-winhttp-baseline.json", implementation: "Raw WinHttpRequest"},
		{file: "vba-web-baseline.json", implementation: "VBA-Web", expectedCommit: "cefc320acc5372e0b86eed1d20eb3f31b331d598"},
	}

	for _, test := range tests {
		t.Run(test.implementation, func(t *testing.T) {
			content, err := os.ReadFile(filepath.Join("..", "..", "benchmarks", "results", test.file))
			if err != nil {
				t.Fatal(err)
			}
			var result struct {
				SchemaVersion  int    `json:"schema_version"`
				Benchmark      string `json:"benchmark"`
				Implementation struct {
					Name         string `json:"name"`
					SourceCommit string `json:"source_commit"`
				} `json:"implementation"`
				Server struct {
					ExternalNetwork bool `json:"external_network"`
				} `json:"server"`
				Parameters struct {
					WarmupIterations  int   `json:"warmup_iterations"`
					LatencyIterations int   `json:"latency_iterations"`
					DownloadBytes     int64 `json:"download_bytes"`
				} `json:"parameters"`
				Results []struct {
					Scenario string  `json:"scenario"`
					Status   int     `json:"status"`
					Elapsed  float64 `json:"elapsed_ms"`
					Bytes    int64   `json:"bytes"`
				} `json:"results"`
			}
			if err := json.Unmarshal(content, &result); err != nil {
				t.Fatal(err)
			}
			if result.SchemaVersion != 1 || result.Benchmark != "http-client" || result.Implementation.Name != test.implementation || result.Implementation.SourceCommit != test.expectedCommit {
				t.Fatalf("unexpected benchmark identity: %#v", result)
			}
			if result.Server.ExternalNetwork || result.Parameters.WarmupIterations != 5 || result.Parameters.LatencyIterations != 50 || result.Parameters.DownloadBytes < 100*1024*1024 {
				t.Fatalf("unexpected benchmark parameters: %#v", result.Parameters)
			}
			if len(result.Results) != 2 {
				t.Fatalf("result count = %d, want 2", len(result.Results))
			}
			if result.Results[0].Scenario != "sequential_get" || result.Results[0].Status != 204 || result.Results[0].Elapsed <= 0 {
				t.Fatalf("unexpected latency result: %#v", result.Results[0])
			}
			if result.Results[1].Scenario != "buffered_download" || result.Results[1].Status != 200 || result.Results[1].Bytes != 100*1024*1024 || result.Results[1].Elapsed <= 0 {
				t.Fatalf("unexpected download result: %#v", result.Results[1])
			}
		})
	}
}
