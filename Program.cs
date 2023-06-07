using DocumentEditorCore;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowAllOrigins", builder =>
    {
        builder.AllowAnyOrigin()
        .AllowAnyMethod()
        .AllowAnyHeader();
    });
});
builder.Services.AddMvcCore()
        .AddNewtonsoftJson();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();
app.UseCors("AllowAllOrigins");


var documentEditorHelper = new DocumentEditorHelper();


app.MapPost("/Import", (HttpRequest request) =>
{

    return documentEditorHelper.Import(request.Form);

});

app.MapPost("/ExportSFDT", (SaveParameter data) =>
{
    Microsoft.AspNetCore.Mvc.FileStreamResult fileStream = documentEditorHelper.ExportSFDT(data);
    return Results.File(fileStream.FileStream, fileStream.ContentType, fileStream.FileDownloadName);
});

app.MapPost("/ExportSFDTtoPDF", (SaveParameter data) =>
{
    Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("ORg4AjUWIQA/Gnt2VFhiQllPd11dXmJWd1p/THNYflR1fV9DaUwxOX1dQl9gSXhTcERjWXZfdXRUT2Y=");

    Microsoft.AspNetCore.Mvc.FileStreamResult fileStream = documentEditorHelper.ExportSFDTtoPDF(data);
    return Results.File(fileStream.FileStream, fileStream.ContentType, fileStream.FileDownloadName);
});
app.MapGet("/", () =>
{
    
    return "Resumas syncfusion service is running";
});

Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("ORg4AjUWIQA/Gnt2VFhiQllPd11dXmJWd1p/THNYflR1fV9DaUwxOX1dQl9gSXhTcERjWXZfdXRUT2Y=");

app.Run("http://0.0.0.0:80");
