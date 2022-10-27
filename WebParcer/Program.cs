using Microsoft.EntityFrameworkCore;
using WebParcer.DBContext;
using WebParcer.Services;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddMvc();                          
builder.Services.AddDbContext<ApplicationDBContext>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("SQLConnection")));      //Подключение базы данных
builder.Services.AddScoped<ExcelParcingService>();                                          //Добавление сервиса парсинга файлов

var app = builder.Build();

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();

app.UseRouting();

app.UseAuthorization();

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=Index}/{id?}");

app.Run();
