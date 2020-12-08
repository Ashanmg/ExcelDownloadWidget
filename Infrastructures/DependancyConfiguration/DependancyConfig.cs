using Autofac;
using AutoMapper;
using ExcelDownloadWidget.Mappings;
using ExcelDownloadWidget.Repository;
using ExcelDownloadWidget.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExcelDownloadWidget.Infrastructures.DependancyConfiguration
{
    public class DependancyConfig
    {
        public void Load(ContainerBuilder builder)
        {
            builder
                .RegisterAssemblyTypes(typeof(IRepository).Assembly)
                .Where(type => typeof(IRepository).IsAssignableFrom(type))
                .InstancePerLifetimeScope()
                .AsImplementedInterfaces();

            builder
                .RegisterAssemblyTypes(typeof(IService).Assembly)
                .Where(type => typeof(IService).IsAssignableFrom(type))
                .InstancePerLifetimeScope()
                .AsImplementedInterfaces();

            RegisterMappings(builder);
        }

        private void RegisterMappings(ContainerBuilder builder)
        {
            var mappingProfiles = typeof(IMapping).Assembly
                .GetTypes()
                .Where(type => typeof(Profile).IsAssignableFrom(type))
                .Select(p => (Profile)Activator.CreateInstance(p));

            builder.Register(ctx => new MapperConfiguration(cfg =>
            {
                foreach (var profile in mappingProfiles)
                {
                    cfg.AddProfile(profile);
                }
            }));

            builder
                .Register(ctx => ctx.Resolve<MapperConfiguration>().CreateMapper())
                .As<IMapper>()
                .InstancePerLifetimeScope();
        }
    }
}