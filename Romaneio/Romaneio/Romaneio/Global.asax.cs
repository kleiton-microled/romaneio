using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Optimization;
using System.Web.Routing;
using System.Web.Mvc;
using Unity;
using Unity.Lifetime;
using Romaneio.Interfaces;
using Romaneio.Models;
using Romaneio.Classes;
using Romaneio.Repositorios;
using Unity.AspNet.Mvc;

namespace Romaneio
{
    public class MvcApplication : System.Web.HttpApplication
    {
        protected void Application_Start()
        {
            AreaRegistration.RegisterAllAreas();
            FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters);
            RouteConfig.RegisterRoutes(RouteTable.Routes);
            BundleConfig.RegisterBundles(BundleTable.Bundles);

            var container = new UnityContainer();

            // Registrando dependências
            container.RegisterType<IDadosRepositorio,  DadosRepositorio>(new HierarchicalLifetimeManager());
            container.RegisterType<IFechamentosRepositorio, FechamentosRepositorio>(new HierarchicalLifetimeManager());
            container.RegisterType<IManutencaoCLARepositorio, ManutencaoCLARepositorio>(new HierarchicalLifetimeManager());
            container.RegisterType<IMovimentacaoCSRepositorio, MovimentacaoCSRepositorio>(new HierarchicalLifetimeManager()); 
            container.RegisterType<IRomaneioRepositorio, RomaneioRepositorio>(new HierarchicalLifetimeManager());
            container.RegisterType<ICompartilhaScannerRepositorio, CompartilhaScannerRepositorio>(new HierarchicalLifetimeManager());
            container.RegisterType<IAssociacaoMarcantesRepositorio, AssociacaoMarcantesRepositorio>(new HierarchicalLifetimeManager());
            container.RegisterType<IConsultaLiberacaoCarregamentoRepositorio, ConsultaLiberacaoCarregamentoRepositorio>(new HierarchicalLifetimeManager());
            container.RegisterType<ICarregamentoCargaSoltaRepositorio, CarregamentoCargaSoltaRepositorio>(new HierarchicalLifetimeManager());
            container.RegisterType<ILocalizacaoCarregamentosRepositorio, LocalizacaoCarregamentosRepositorio>(new HierarchicalLifetimeManager());

            // Registrando o controller
            DependencyResolver.SetResolver(new UnityDependencyResolver(container));


        }
    }
}
