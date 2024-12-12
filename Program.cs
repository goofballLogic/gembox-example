
GemBox.Document.ComponentInfo.SetLicense("FREE-LIMITED-KEY");

var renderer = new WordRenderer();
renderer.AddContentText([
    new ContentText("This should be Bold text", IsBold: true),
]);
renderer.AddContentText([
    new ContentText("This should NOT be bold text")
]);
renderer.AddContentSectionBreak();
renderer.AddContentText([
    new ContentText("This is after a section break (there should be a line above)")
]);

var rendered = renderer.GenerateDocAsPDF();

using var fs = new FileStream("output.pdf", FileMode.Create);
rendered.CopyTo(fs);
fs.Flush();
fs.Close();