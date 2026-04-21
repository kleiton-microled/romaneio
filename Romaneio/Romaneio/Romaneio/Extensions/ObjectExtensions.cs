using System;

namespace Romaneio.Extensions
{
    public static class ObjectExtensions
    {
        public static int ToInt(this object valor)
        {
            if (valor == null)
                return 0;

            if (Int32.TryParse(valor.ToString(), out _))
                return Convert.ToInt32(valor);

            return 0;
        }
    }
}
