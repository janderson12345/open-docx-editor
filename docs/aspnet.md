# Using open-docx-editor in ASP.NET (MVC, Blazor, Web Forms)

The editor runs entirely in the browser. Your server only stores and returns `.docx` bytes, so
any ASP.NET stack can host it: ASP.NET Core MVC or Razor Pages, Blazor (Server or
WebAssembly), and classic ASP.NET Web Forms (`.aspx`) on .NET Framework.

In every case the pattern is the same:

1. **Serve the two files.** Put `dist/open-docx-editor.js` and `dist/open-docx-editor.css`
   (from `npm run build`) in your static files folder, for example
   `wwwroot/lib/open-docx-editor/` (Core) or `/Scripts/open-docx-editor/` (Web Forms).
2. **Add a `<div>`** with a height for the editor to fill.
3. **Load the document:** `await editor.open(await fetch(url), 'name.docx')`.
4. **Save it:** `editor.save()` returns the edited file as bytes (`Uint8Array`), which you
   POST back to the server.

The bundle is a JavaScript module, so load it with `<script type="module">` or `import()`.
That works in every current browser (Edge, Chrome, Firefox, Safari), but not in Internet
Explorer or Edge's IE mode.

```js
const editor = mountDocxEditor(div, {
  onSave: (bytes, name) => upload(bytes),  // the toolbar's Save button; without it, Save downloads the file
  onChange: () => { /* document edited; editor.dirty is now true */ },
});
await editor.open(bytesOrResponseOrFile, 'name.docx');
editor.save();        // Uint8Array of the .docx
editor.dirty;         // edited since open (call editor.markSaved() after a successful upload)
editor.destroy();     // remove it (Blazor component disposal, SPA navigation)
```

---

## ASP.NET Core MVC

**Controller** (`Controllers/DocumentsController.cs`). Replace the file-system calls with your
own storage (database, blob storage, SharePoint-free file share, and so on).

```csharp
using Microsoft.AspNetCore.Mvc;

public class DocumentsController : Controller
{
    private const string DocxType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    private readonly string _root;
    public DocumentsController(IWebHostEnvironment env) => _root = Path.Combine(env.ContentRootPath, "App_Data", "docs");

    // The page that hosts the editor: /Documents/Edit/5
    public IActionResult Edit(int id) => View(id);

    // The document bytes: GET /Documents/File/5
    [HttpGet]
    public IActionResult File(int id)
    {
        var path = Path.Combine(_root, $"{id}.docx");
        if (!System.IO.File.Exists(path)) return NotFound();
        return PhysicalFile(path, DocxType, $"{id}.docx");
    }

    // Save: POST /Documents/File/5 with the .docx as the request body
    [HttpPost, ValidateAntiForgeryToken]
    [RequestSizeLimit(50_000_000)]
    public async Task<IActionResult> File(int id, CancellationToken ct)
    {
        Directory.CreateDirectory(_root);
        await using var fs = System.IO.File.Create(Path.Combine(_root, $"{id}.docx"));
        await Request.Body.CopyToAsync(fs, ct);
        return NoContent();
    }
}
```

**View** (`Views/Documents/Edit.cshtml`):

```cshtml
@model int
@inject Microsoft.AspNetCore.Antiforgery.IAntiforgery Antiforgery
@{
    ViewData["Title"] = "Edit document";
    var token = Antiforgery.GetAndStoreTokens(Context).RequestToken;
}
<link rel="stylesheet" href="~/lib/open-docx-editor/open-docx-editor.css" />

<div id="editor" style="height: calc(100vh - 120px)"></div>

<script type="module">
    import { mountDocxEditor } from '@Url.Content("~/lib/open-docx-editor/open-docx-editor.js")';

    const fileUrl = '@Url.Action("File", "Documents", new { id = Model })';
    const editor = mountDocxEditor(document.getElementById('editor'), {
        async onSave(bytes) {
            const res = await fetch(fileUrl, {
                method: 'POST',
                headers: { 'Content-Type': 'application/octet-stream', 'RequestVerificationToken': '@token' },
                body: bytes,
            });
            if (!res.ok) { alert('Save failed: ' + res.status); return; }
            editor.markSaved();
        },
    });
    await editor.open(await fetch(fileUrl), '@(Model).docx');

    window.addEventListener('beforeunload', (e) => { if (editor.dirty) e.preventDefault(); });
</script>
```

The `RequestVerificationToken` header satisfies `[ValidateAntiForgeryToken]` with the default
antiforgery settings. Razor Pages work the same way, with the two actions as page handlers.

---

## Blazor (Server or WebAssembly)

Blazor talks to the editor through a small JavaScript module. Bytes travel as streams, which
avoids Blazor Server's 32 KB SignalR message limit, so large documents are fine.

**`wwwroot/js/docxEditor.js`:**

```js
import { mountDocxEditor } from '../lib/open-docx-editor/open-docx-editor.js';

export function mount(element, dotnet) {
  const editor = mountDocxEditor(element, {
    // Toolbar Save button: tell the component; it pulls the bytes with getBytes().
    onSave: () => dotnet.invokeMethodAsync('OnSaveRequested'),
    onChange: () => dotnet.invokeMethodAsync('NotifyChanged'),
  });
  return {
    async open(streamRef, name) {
      const bytes = new Uint8Array(await streamRef.arrayBuffer());
      await editor.open(bytes, name);
    },
    getBytes() { return editor.save(); },  // arrives in .NET as IJSStreamReference
    markSaved() { editor.markSaved(); },
    destroy() { editor.destroy(); },
  };
}
```

**`Components/DocxEditor.razor`.** It's a reusable component: give it bytes, and it raises
`OnSave` with the edited bytes.

```razor
@implements IAsyncDisposable
@inject IJSRuntime JS

<link rel="stylesheet" href="lib/open-docx-editor/open-docx-editor.css" />
<div @ref="_host" style="height: @Height"></div>

@code {
    [Parameter, EditorRequired] public byte[] Document { get; set; } = default!;
    [Parameter] public string FileName { get; set; } = "document.docx";
    [Parameter] public string Height { get; set; } = "80vh";
    [Parameter] public EventCallback<byte[]> OnSave { get; set; }
    [Parameter] public EventCallback OnChanged { get; set; }

    private ElementReference _host;
    private IJSObjectReference? _module, _editor;
    private DotNetObjectReference<DocxEditor>? _self;

    protected override async Task OnAfterRenderAsync(bool firstRender)
    {
        if (!firstRender) return;
        _module = await JS.InvokeAsync<IJSObjectReference>("import", "./js/docxEditor.js");
        _self = DotNetObjectReference.Create(this);
        _editor = await _module.InvokeAsync<IJSObjectReference>("mount", _host, _self);
        using var stream = new MemoryStream(Document);
        using var streamRef = new DotNetStreamReference(stream);
        await _editor.InvokeVoidAsync("open", streamRef, FileName);
    }

    /// <summary>The edited document as .docx bytes.</summary>
    public async Task<byte[]> GetBytesAsync()
    {
        var js = await _editor!.InvokeAsync<IJSStreamReference>("getBytes");
        await using var s = await js.OpenReadStreamAsync(maxAllowedSize: 100_000_000);
        using var ms = new MemoryStream();
        await s.CopyToAsync(ms);
        return ms.ToArray();
    }

    [JSInvokable] public async Task OnSaveRequested()
    {
        await OnSave.InvokeAsync(await GetBytesAsync());
        await _editor!.InvokeVoidAsync("markSaved");
    }

    [JSInvokable] public Task NotifyChanged() => OnChanged.InvokeAsync();

    public async ValueTask DisposeAsync()
    {
        try
        {
            if (_editor is not null) { await _editor.InvokeVoidAsync("destroy"); await _editor.DisposeAsync(); }
            if (_module is not null) await _module.DisposeAsync();
        }
        catch (JSDisconnectedException) { } // circuit already gone (Blazor Server)
        _self?.Dispose();
    }
}
```

**Using it on a page:**

```razor
@page "/documents/{Id:int}"
@inject IDocumentStore Store   @* your own service *@

@if (_bytes is null) { <p>Loading…</p> }
else
{
    <DocxEditor Document="_bytes" FileName="@($"{Id}.docx")" OnSave="SaveAsync" />
}

@code {
    [Parameter] public int Id { get; set; }
    private byte[]? _bytes;

    protected override async Task OnParametersSetAsync() => _bytes = await Store.LoadAsync(Id);

    private Task SaveAsync(byte[] docx) => Store.SaveAsync(Id, docx);
}
```

Add `@rendermode InteractiveServer` (or `InteractiveWebAssembly`) to the page on .NET 8+
if the app doesn't set an interactive render mode globally. The component needs interactivity
to run JavaScript.

---

## ASP.NET Web Forms (`.aspx`, .NET Framework)

Yes, this works too. The editor is plain HTML and JavaScript, so a Web Forms page only needs
to host the `<div>` and the script, plus an HTTP handler that reads and writes the bytes.

**`DocumentHandler.ashx`:**

```csharp
<%@ WebHandler Language="C#" Class="DocumentHandler" %>
using System.IO;
using System.Web;

public class DocumentHandler : IHttpHandler
{
    private const string DocxType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    public bool IsReusable => true;

    public void ProcessRequest(HttpContext context)
    {
        int id;
        if (!int.TryParse(context.Request.QueryString["id"], out id)) { context.Response.StatusCode = 400; return; }
        var path = context.Server.MapPath("~/App_Data/docs/" + id + ".docx");

        if (context.Request.HttpMethod == "GET")
        {
            if (!File.Exists(path)) { context.Response.StatusCode = 404; return; }
            context.Response.ContentType = DocxType;
            context.Response.TransmitFile(path);
        }
        else if (context.Request.HttpMethod == "POST")
        {
            // Check the user may edit this document here (the handler skips page authorization logic).
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            using (var fs = File.Create(path)) context.Request.InputStream.CopyTo(fs);
            context.Response.StatusCode = 204;
        }
        else context.Response.StatusCode = 405;
    }
}
```

**`EditDocument.aspx`:**

```aspx
<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="EditDocument.aspx.cs" Inherits="MyApp.EditDocument" %>
<!DOCTYPE html>
<html>
<head runat="server">
    <title>Edit document</title>
    <link rel="stylesheet" href="<%= ResolveUrl("~/Scripts/open-docx-editor/open-docx-editor.css") %>" />
</head>
<body>
    <form id="form1" runat="server">
        <div id="editor" style="height: 90vh"></div>
    </form>

    <script type="module">
        import { mountDocxEditor } from '<%= ResolveUrl("~/Scripts/open-docx-editor/open-docx-editor.js") %>';

        const fileUrl = '<%= ResolveUrl("~/DocumentHandler.ashx") %>?id=<%= Request.QueryString["id"] %>';
        const editor = mountDocxEditor(document.getElementById('editor'), {
            async onSave(bytes) {
                const res = await fetch(fileUrl, { method: 'POST', body: bytes, headers: { 'Content-Type': 'application/octet-stream' } });
                if (res.ok) editor.markSaved(); else alert('Save failed: ' + res.status);
            },
        });
        await editor.open(await fetch(fileUrl), 'document.docx');
    </script>
</body>
</html>
```

Notes for Web Forms:

- The editor's toolbar buttons are `type="button"`, so they don't trigger a postback inside
  `<form runat="server">`.
- For large documents, raise `maxRequestLength` (KB) and `maxAllowedContentLength` (bytes) in
  `web.config`.
- To save as part of an ordinary postback instead of the handler, copy the bytes into a hidden
  field before submitting: `hidden.value = btoa(String.fromCharCode(...bytes))` for small files,
  or chunked base64 for big ones. Then call `Convert.FromBase64String(HiddenDocx.Value)` in the
  code-behind. The `fetch` handler above is simpler and handles large files better.
- Validate the `id` against the signed-in user before reading or writing. Handlers don't run
  page-level authorization code for you.

---

## Common questions

- **Where do the bytes come from?** Anywhere: a database `varbinary(max)`, Azure Blob Storage,
  a file share, or an upload. The editor only needs the `.docx` bytes, and it gives back new
  bytes that keep every part of the original file it didn't change.
- **Can it open files from the user's PC?** Yes. The toolbar's Open… button, or
  `editor.open(fileInput.files[0])`. On managed Windows PCs the browser or antivirus may
  scan a picked file before handing it over, which can take a few seconds. That delay is
  outside the editor.
- **Read-only viewing?** For now, set `editor.view.setProps({ editable: () => false })` after
  opening, and hide the toolbar with CSS (`.odx-toolbar { display: none }`).
