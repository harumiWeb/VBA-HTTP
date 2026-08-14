# Examples

The snippets below assume the production modules are installed in the current
workbook and that `VBAHttp` is available.

## GET JSON/text

```vb
Public Sub GetJson()
    Dim client As HttpClient
    Dim response As HttpResponse

    Set client = VBAHttp.CreateClient()
    Set response = client.GetResponse("https://api.example.test/items")
    response.RaiseForStatus
    Debug.Print response.Text
End Sub
```

## Request with headers, query, and a text body

```vb
Public Sub PostJson()
    Dim request As HttpRequest
    Dim response As HttpResponse

    Set request = VBAHttp.CreateRequest()
    request.Method = "POST"
    request.Url = "https://api.example.test/items"
    request.Query.Add "dry_run", True
    request.Headers.SetValue "Accept", "application/json"
    request.Headers.SetValue "Content-Type", "application/json; charset=utf-8"
    request.Body.Text = "{""name"":""sample""}"

    Set response = VBAHttp.CreateClient().Execute(request)
    response.RaiseForStatus
End Sub
```

## Binary response

```vb
Dim bytes As Variant
Set response = client.GetResponse("https://files.example.test/icon.bin")
response.RaiseForStatus
bytes = response.Body.Bytes
Debug.Print UBound(bytes) - LBound(bytes) + 1
```

Use `DownloadFile` instead when the payload should not be buffered in memory.

## Retry and deadline

```vb
Dim policy As HttpRetryPolicy
Dim options As HttpExecutionOptions

Set policy = VBAHttp.CreateRetryPolicy()
policy.MaxAttempts = 5
policy.BaseDelayMilliseconds = 100
policy.MaxDelayMilliseconds = 2000
Set options = VBAHttp.CreateExecutionOptions()
Set options.RetryPolicy = policy
options.TotalDeadlineMilliseconds = 15000

Set response = client.GetResponse("https://api.example.test/slow", options)
response.RaiseForStatus
```

## Bounded batch

```vb
Dim urls As New Collection
Dim options As HttpBatchOptions
Dim result As HttpBatchResult
Dim i As Long

urls.Add "https://api.example.test/a"
urls.Add "https://api.example.test/b"
urls.Add "https://api.example.test/c"
Set options = VBAHttp.CreateBatchOptions()
options.MaxConcurrency = 8
Set result = VBAHttp.CreateClient().GetMany(urls, options)

For i = 1 To result.Count
    Debug.Print result.ItemAt(i).Status
Next
```

## Cancellation

```vb
Dim token As HttpCancellationToken
Dim options As HttpExecutionOptions

Set token = VBAHttp.CreateCancellationToken()
Set options = VBAHttp.CreateExecutionOptions()
Set options.CancellationToken = token
options.TotalDeadlineMilliseconds = 60000

' Keep token in module state when a UI button or scheduled callback will cancel.
' token.Cancel
Set response = client.GetResponse("https://api.example.test/long", options)
```

## Native download with progress

```vb
Dim nativeClient As HttpClient
Dim download As HttpDownloadResult

Set nativeClient = VBAHttp.CreateNativeClient()
Set download = nativeClient.DownloadFile( _
    "https://files.example.test/archive.bin", _
    "C:\Temp\archive.bin")
download.RaiseForStatus
```

Implement `IHttpProgressSink` in a class module and pass the instance as the
fourth argument to `DownloadFile` or the final argument to upload methods.

## Native upload and multipart

```vb
Dim upload As HttpUploadResult
Dim form As HttpMultipartForm

Set upload = nativeClient.UploadFile( _
    "https://files.example.test/upload", _
    "C:\Temp\payload.bin", _
    "application/octet-stream")
upload.RaiseForStatus

Set form = VBAHttp.CreateMultipartForm()
form.AddField "description", "sample"
form.AddFile "payload", "C:\Temp\payload.bin"
Set upload = nativeClient.UploadMultipart("https://files.example.test/form", form)
upload.RaiseForStatus
```

## Basic or Bearer authentication

```vb
Dim auth As IHttpAuthProvider

Set auth = VBAHttp.CreateBearerAuthProvider("opaque-token")
Set client.AuthProvider = auth
Set response = client.GetResponse("https://api.example.test/private")
```

Credentials require HTTPS by default. For a fixed local test server only, pass
`True` as `AllowInsecureHttp`; never use that escape for Internet traffic.

## Cookies and diagnostics

```vb
Dim jar As HttpCookieJar
Dim diagnostics As HttpDiagnostics

Set jar = VBAHttp.CreateCookieJar()
Set diagnostics = VBAHttp.CreateDiagnostics()
diagnostics.Enabled = True
Set client.CookieJar = jar
Set client.Diagnostics = diagnostics

Set response = client.GetResponse("https://api.example.test/session")
Debug.Print jar.Count
Debug.Print diagnostics.ToJson
```

Cookie state is opt-in and in-memory. Diagnostic output redacts cookie,
authorization, and proxy-authorization values.

## Native HTTP/2 fallback

```vb
Dim protocols As HttpProtocolOptions
Set protocols = VBAHttp.CreateProtocolOptions()
protocols.AllowHttp2 = True
protocols.Mode = HttpProtocolAllowFallback
Set nativeClient.ProtocolOptions = protocols
Set response = nativeClient.GetResponse("https://api.example.test/ping")
Debug.Print response.ProtocolUsed
```

Do not set `AllowHttp3` for production use. HTTP/3/QUIC is unsupported by the
current compatibility policy.
