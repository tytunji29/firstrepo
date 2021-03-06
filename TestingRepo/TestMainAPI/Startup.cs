using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.HttpsPolicy;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Swashbuckle.AspNetCore.Swagger;
using TestAPI.Filters;
using TestingRepo.Repository;

namespace TestMainAPI
{
    public class Startup
    {
        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;
        }

        public IConfiguration Configuration { get; }

        // This method gets called by the runtime. Use this method to add services to the container.
        public void ConfigureServices(IServiceCollection services)
        {
            try
            {
                
                #region Swagger Config
                services.AddSwaggerGen(c =>
                {
                    c.SwaggerDoc("v1", new Info { Title = "TBOSS", Version = "v1" });
                });

                services.ConfigureSwaggerGen(options =>
                {
                    options.OperationFilter<SwaggerAuthorizationHeaderParameterOperationFilter>();
                });
                #endregion


                services.AddScoped<ICountryRepository, CountryRepository>();
            }
            catch (Exception)
            {
                throw;
            }
            
        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app, IHostEnvironment env)
        {
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }
            else
            {
                //app.UseWebApiExceptionHandler();
                // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
                app.UseHsts();
            }
            app.UseSwagger();
            app.UseSwaggerUI(c =>
            {
                c.SwaggerEndpoint("/swagger/v1/swagger.json", "Ty TEST");
                c.RoutePrefix = string.Empty;
            });


            app.UseAuthentication();
            app.UseHttpsRedirection();
            app.UseRouting();
           // app.UseAuthorization();
            app.UseWelcomePage();

        }
    }
}
