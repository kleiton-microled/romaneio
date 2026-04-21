using System;

namespace Romaneio.Extensions
{
    public static class BooleanExtensions
    {
        public static int ToInt(this bool valor)
            => Convert.ToInt32(valor);

        public static string ToSimOuNao(this bool valor)
            => valor ? "Sim" : "Não";
    }
}
