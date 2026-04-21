using System.Security.Claims;
using System.Security.Principal;

namespace Romaneio.Extensions
{
    public static class UserExtensions
    {
        public static int ObterId(this IPrincipal user)
        {
            Claim claim = ((ClaimsIdentity)user.Identity).FindFirst("Id");

            return claim == null
                ? 0
                : claim.Value.ToInt();
        }

        public static string ObterNome(this IPrincipal user)
        {
            Claim claim = ((ClaimsIdentity)user.Identity).FindFirst("Nome");

            return claim?.Value;
        }


        public static string ObterCPF(this IPrincipal user)
        {
            Claim claim = ((ClaimsIdentity)user.Identity).FindFirst("CPF");

            return claim?.Value;
        }

        public static string ObterEmail(this IPrincipal user)
        {
            Claim claim = ((ClaimsIdentity)user.Identity).FindFirst("Email");

            return claim?.Value;
        }

        public static int ObterEmpresa(this IPrincipal user)
        {
            Claim claim = ((ClaimsIdentity)user.Identity).FindFirst("EmpresaId");

            return claim?.Value != null
                ? claim.Value.ToInt()
                : 0;
        }

        public static string ObterEmpresaDescricao(this IPrincipal user)
        {
            Claim claim = ((ClaimsIdentity)user.Identity).FindFirst("EmpresaDescricao");

            return claim?.Value;
        }

        public static string ObterEmpresaCNPJ(this IPrincipal user)
        {
            Claim claim = ((ClaimsIdentity)user.Identity).FindFirst("EmpresaCNPJ");

            return claim?.Value;
        }

        public static int[] ObterPatiosId(this IPrincipal user)
        {
            Claim claim = ((ClaimsIdentity)user.Identity).FindFirst("EmpresaLogin");

            int[] patios = claim?.Value.ToInt() == 1
                ? new int[] { 1, 7, 9 }
                : new int[] { 3, 4, 5 };

            return patios;
        }

        public static int ObterEmpresaLogin(this IPrincipal user)
        {
            Claim claim = ((ClaimsIdentity)user.Identity).FindFirst("EmpresaLogin");

            return claim?.Value != null
                ? claim.Value.ToInt()
                : 1;
        }

        public static int ObterGrupoId(this IPrincipal user)
        {
            Claim claim = ((ClaimsIdentity)user.Identity).FindFirst("GrupoId");

            return claim?.Value != null
                ? claim.Value.ToInt()
                : 0;
        }

        public static bool UsuarioAdministrador(this IPrincipal user)
        {
            return user.IsInRole("Administrador");
           
        }

        public static bool UsuarioCliente(this IPrincipal user)
        {
            return user.IsInRole("Cliente");

        }

        public static bool UsuarioAdminCliente(this IPrincipal user)
        {
            return user.IsInRole("AdminCliente");

        }

        public static bool UsuarioOperacional(this IPrincipal user)
        {
            return user.IsInRole("Operacional");
        }

        public static bool AcessoExterno(this IPrincipal user)
        {
            return user.IsInRole("AcessoExterno");
        }

        public static bool UsuarioDespachante(this IPrincipal user)
        {
            return user.IsInRole("Despachante");
        }

        public static bool UsuarioImportador(this IPrincipal user)
        {
            return user.IsInRole("Importador");
        }

        public static bool UsuarioTransportadora(this IPrincipal user)
        {
            return user.IsInRole("Transportadora");
        }

        public static bool UsuarioDespachanteRestrito(this IPrincipal user)
        {
            return user.IsInRole("DespachanteRestrito");
        }

        public static bool UsuarioNVOCC(this IPrincipal user)
        {
            return user.IsInRole("NVOCC");
        }

        public static bool VerComoCliente(this IPrincipal user)
        {
            Claim claim = ((ClaimsIdentity)user.Identity).FindFirst("VerComoCliente");

            return claim?.Value.ToInt() > 0
                ? true
                : false;
        }

        public static bool Autenticado(this IPrincipal user)
        {
            return user.Identity.IsAuthenticated && user.ObterEmpresa() > 0;
        }

        public static bool AcessoLiberado(this IPrincipal user)
        {
            Claim claim = ((ClaimsIdentity)user.Identity).FindFirst("Ativo");

            return claim?.Value == "true"
                ? true
                : false;
        }

        public static bool AcessoAgendamentoLTL(this IPrincipal user)
        {
            Claim claim = ((ClaimsIdentity)user.Identity).FindFirst("AcessoAgendamentoLTL");

            return claim?.Value.ToInt() > 0
                ? true
                : false;
        }
    }
}