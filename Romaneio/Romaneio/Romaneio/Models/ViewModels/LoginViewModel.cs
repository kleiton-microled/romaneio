using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace Romaneio.Models
{
    public class LoginViewModel
    {
        public string nome { get; set; }

        [Required(ErrorMessage = "Campo obrigatório")]
        public string Login { get; set; }

        [Display(Name = "Senha")]
        [Required(ErrorMessage = "Campo obrigatório")]
        public string Senha { get; set; }

        public string ReturnUrl { get; set; }

        public bool SenhaValida(string senha)
        {
            return Senha == senha;
        }
      
        [Required(ErrorMessage = "Campo obrigatório")]
        public int PATIO  { get; set; }
        public List<TiposPatios> TIPOS_PATIOS { get; set; }
    }
}