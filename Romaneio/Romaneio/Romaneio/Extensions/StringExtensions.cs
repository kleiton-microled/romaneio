using System;
using System.Collections.Generic;
using System.Data.SqlTypes;
using System.Text;
using System.Text.RegularExpressions;

namespace Romaneio.Extensions
{
    public static class StringExtensions
    {
        public static int ToInt(this string valor)
        {
            if (Int32.TryParse(valor, out _))
                return Convert.ToInt32(valor);

            return 0;
        }

        public static decimal ToDecimal(this string valor)
        {
            if (Decimal.TryParse(valor, out _))
                return Convert.ToDecimal(valor);

            return 0;
        }

        public static double ToDouble(this string valor)
        {
            if (Double.TryParse(valor, out _))
                return Convert.ToDouble(valor);

            return 0;
        }

        public static bool ToBoolean(this string valor)
        {
            if (Boolean.TryParse(valor, out _))
                return true;

            return false;
        }

        public static DateTime ToDateTime(this string valor)
        {
            if (DateTime.TryParse(valor, out DateTime resultado))
                return resultado;

            return SqlDateTime.MinValue.Value;
        }

        public static string ToNumero(this string texto)
        {
            if (string.IsNullOrEmpty(texto))
                texto = "0";

            if (Decimal.TryParse(texto, out _))
                return string.Format("{0:N2}", Convert.ToDecimal(texto));

            return string.Empty;
        }

        public static byte[] ToByteArray(this string texto)
        {
            return Encoding.ASCII.GetBytes(texto);
        }

        public static string RemoverCaracteresEspeciais(this string valor)
        {
            return Regex.Replace(valor ?? string.Empty, "[^0-9a-zA-Z]+", "");
        }

        public static string LimitarTexto(this string valor, int tamanho)
        {
            if (!string.IsNullOrEmpty(valor))
            {
                if (valor.Length > tamanho)
                {
                    return string.Concat(valor.Substring(0, tamanho), "...");
                }
            }

            return valor;
        }

        public static string RemoverTagsHtml(this string valor)
        {
            var tags = new List<string>()
            {
                "<a>",
                "<abbr>",
                "<acronym>",
                "<address>",
                "<applet>",
                "<area>",
                "<article>",
                "<aside>",
                "<audio>",
                "<b>",
                "<base>",
                "<basefont>",
                "<bdi>",
                "<bdo>",
                "<big>",
                "<blockquote>",
                "<body>",
                "<br>",
                "<br/>",
                "<br />",
                "<button>",
                "<canvas>",
                "<caption>",
                "<center>",
                "<cite>",
                "<code>",
                "<col>",
                "<colgroup>",
                "<command>",
                "<datalist>",
                "<dd>",
                "<del>",
                "<details>",
                "<dfn>",
                "<dir>",
                "<div>",
                "<dl>",
                "<dt>",
                "<em>",
                "<embed>",
                "<fieldset>",
                "<figcaption>",
                "<figure>",
                "<font>",
                "<footer>",
                "<form>",
                "<frame>",
                "<frameset>",
                "<h1>",
                "<h2>",
                "<h3>",
                "<h4>",
                "<h5>",
                "<h6>",
                "<head>",
                "<header>",
                "<hgroup>",
                "<hr>",
                "<html>",
                "<i>",
                "<iframe>",
                "<img>",
                "<input>",
                "<ins>",
                "<kbd>",
                "<keygen>",
                "<label>",
                "<legend>",
                "<li>",
                "<link>",
                "<map>",
                "<mark>",
                "<menu>",
                "<meta>",
                "<meter>",
                "<nav>",
                "<noframes>",
                "<noscript>",
                "<object>",
                "<ol>",
                "<optgroup>",
                "<option>",
                "<output>",
                "<p>",
                "<param>",
                "<pre>",
                "<progress>",
                "<q>",
                "<rp>",
                "<rt>",
                "<ruby>",
                "<s>",
                "<samp>",
                "<script>",
                "<section>",
                "<select>",
                "<small>",
                "<source>",
                "<span>",
                "<strike>",
                "<strong>",
                "<style>",
                "<sub>",
                "<summary>",
                "<sup>",
                "<table>",
                "<tbody>",
                "<td>",
                "<textarea>",
                "<tfoot>",
                "<th>",
                "<thead>",
                "<time>",
                "<title>",
                "<tr>",
                "<track>",
                "<tt>",
                "<u>",
                "<ul>",
                "<var>",
                "<video>",
                "<wbr>",
                "</a>",
                "</abbr>",
                "</acronym>",
                "</address>",
                "</applet>",
                "</area>",
                "</article>",
                "</aside>",
                "</audio>",
                "</b>",
                "</base>",
                "</basefont>",
                "</bdi>",
                "</bdo>",
                "</big>",
                "</blockquote>",
                "</body>",
                "</br>",
                "</button>",
                "</canvas>",
                "</caption>",
                "</center>",
                "</cite>",
                "</code>",
                "</col>",
                "</colgroup>",
                "</command>",
                "</datalist>",
                "</dd>",
                "</del>",
                "</details>",
                "</dfn>",
                "</dir>",
                "</div>",
                "</dl>",
                "</dt>",
                "</em>",
                "</embed>",
                "</fieldset>",
                "</figcaption>",
                "</figure>",
                "</font>",
                "</footer>",
                "</form>",
                "</frame>",
                "</frameset>",
                "</h1>",
                "</h2>",
                "</h3>",
                "</h4>",
                "</h5>",
                "</h6>",
                "</head>",
                "</header>",
                "</hgroup>",
                "</hr>",
                "</html>",
                "</i>",
                "</iframe>",
                "</img>",
                "</input>",
                "</ins>",
                "</kbd>",
                "</keygen>",
                "</label>",
                "</legend>",
                "</li>",
                "</link>",
                "</map>",
                "</mark>",
                "</menu>",
                "</meta>",
                "</meter>",
                "</nav>",
                "</noframes>",
                "</noscript>",
                "</object>",
                "</ol>",
                "</optgroup>",
                "</option>",
                "</output>",
                "</p>",
                "</param>",
                "</pre>",
                "</progress>",
                "</q>",
                "</rp>",
                "</rt>",
                "</ruby>",
                "</s>",
                "</samp>",
                "</script>",
                "</section>",
                "</select>",
                "</small>",
                "</source>",
                "</span>",
                "</strike>",
                "</strong>",
                "</style>",
                "</sub>",
                "</summary>",
                "</sup>",
                "</table>",
                "</tbody>",
                "</td>",
                "</textarea>",
                "</tfoot>",
                "</th>",
                "</thead>",
                "</time>",
                "</title>",
                "</tr>",
                "</track>",
                "</tt>",
                "</u>",
                "</ul>",
                "</var>",
                "</video>",
                "</wbr>",
                "&nbsp;"
            };

            foreach (var tag in tags)
            {
                valor = valor.Replace(tag, string.Empty);
            }

            return valor;
        }

        public static string EscreverValorPorExtenso(this decimal valor)
        {
            if (valor <= 0)
                return string.Empty;
            else
            {
                string montagem = string.Empty;
                if (valor > 0 & valor < 1)
                {
                    valor *= 100;
                }
                string strValor = valor.ToString("000");
                int a = Convert.ToInt32(strValor.Substring(0, 1));
                int b = Convert.ToInt32(strValor.Substring(1, 1));
                int c = Convert.ToInt32(strValor.Substring(2, 1));
                if (a == 1) montagem += (b + c == 0) ? "Cem" : "Cento";
                else if (a == 2) montagem += "Duzentos";
                else if (a == 3) montagem += "Trezentos";
                else if (a == 4) montagem += "Quatrocentos";
                else if (a == 5) montagem += "Quinhentos";
                else if (a == 6) montagem += "Seiscentos";
                else if (a == 7) montagem += "Setecentos";
                else if (a == 8) montagem += "Oitocentos";
                else if (a == 9) montagem += "Novecentos";
                if (b == 1)
                {
                    if (c == 0) montagem += ((a > 0) ? " e " : string.Empty) + "Dez";
                    else if (c == 1) montagem += ((a > 0) ? " e " : string.Empty) + "Onze";
                    else if (c == 2) montagem += ((a > 0) ? " e " : string.Empty) + "Doze";
                    else if (c == 3) montagem += ((a > 0) ? " e " : string.Empty) + "Treze";
                    else if (c == 4) montagem += ((a > 0) ? " e " : string.Empty) + "Quatorze";
                    else if (c == 5) montagem += ((a > 0) ? " e " : string.Empty) + "Quinze";
                    else if (c == 6) montagem += ((a > 0) ? " e " : string.Empty) + "Dezesseis";
                    else if (c == 7) montagem += ((a > 0) ? " e " : string.Empty) + "Dezessete";
                    else if (c == 8) montagem += ((a > 0) ? " e " : string.Empty) + "Dezoito";
                    else if (c == 9) montagem += ((a > 0) ? " e " : string.Empty) + "Dezenove";
                }
                else if (b == 2) montagem += ((a > 0) ? " e " : string.Empty) + "Vinte";
                else if (b == 3) montagem += ((a > 0) ? " e " : string.Empty) + "Trinta";
                else if (b == 4) montagem += ((a > 0) ? " e " : string.Empty) + "Quarenta";
                else if (b == 5) montagem += ((a > 0) ? " e " : string.Empty) + "Cinquenta";
                else if (b == 6) montagem += ((a > 0) ? " e " : string.Empty) + "Sessenta";
                else if (b == 7) montagem += ((a > 0) ? " e " : string.Empty) + "Setenta";
                else if (b == 8) montagem += ((a > 0) ? " e " : string.Empty) + "Oitenta";
                else if (b == 9) montagem += ((a > 0) ? " e " : string.Empty) + "Noventa";
                if (strValor.Substring(1, 1) != "1" & c != 0 & montagem != string.Empty) montagem += " e ";
                if (strValor.Substring(1, 1) != "1")
                    if (c == 1) montagem += "Um";
                    else if (c == 2) montagem += "Dois";
                    else if (c == 3) montagem += "Três";
                    else if (c == 4) montagem += "Quatro";
                    else if (c == 5) montagem += "Cinco";
                    else if (c == 6) montagem += "Seis";
                    else if (c == 7) montagem += "Sete";
                    else if (c == 8) montagem += "Oito";
                    else if (c == 9) montagem += "Nove";
                return montagem;
            }
        }

        public static string FormatarConteiner(this string valor)
        {
            
           
            if (!string.IsNullOrEmpty(valor))
            {
                if (valor.Length == 11)
                {
                    valor = valor.Substring(0, 10) + "-" + valor.Substring(valor.Length - 1);
                }

                if (valor.Length == 10)
                {
                    valor = valor.Substring(0, 10) + "-_";
                }

                if (valor.Length == 9)
                {
                    valor = valor.Substring(0, 9) + "_-_";
                }

         
            }
            
            return valor;
        }

          public static string ControllerName(this string valor)
          {
            return valor.Substring(0, valor.LastIndexOf("Controller"));
        }


        public static bool IsDate(this string data)
        {
            if (DateTime.TryParse(data, out _))
                return true;

            return false;
        }
        public static string toExtenso(decimal valor)
        {
            if (valor <= 0 | valor >= 1000000000000000)
                return "Valor não suportado pelo sistema.";
            else
            {
                string strValor = valor.ToString("000000000000000.00");
                string valor_por_extenso = string.Empty;

                for (int i = 0; i <= 15; i += 3)
                {
                    valor_por_extenso += escreva_parte(Convert.ToDecimal(strValor.Substring(i, 3)));
                    if (i == 0 & valor_por_extenso != string.Empty)
                    {
                        if (Convert.ToInt32(strValor.Substring(0, 3)) == 1)
                            valor_por_extenso += " TRILHÃO" + ((Convert.ToDecimal(strValor.Substring(3, 12)) > 0) ? " E " : string.Empty);
                        else if (Convert.ToInt32(strValor.Substring(0, 3)) > 1)
                            valor_por_extenso += " TRILHÕES" + ((Convert.ToDecimal(strValor.Substring(3, 12)) > 0) ? " E " : string.Empty);
                    }
                    else if (i == 3 & valor_por_extenso != string.Empty)
                    {
                        if (Convert.ToInt32(strValor.Substring(3, 3)) == 1)
                            valor_por_extenso += " BILHÃO" + ((Convert.ToDecimal(strValor.Substring(6, 9)) > 0) ? " E " : string.Empty);
                        else if (Convert.ToInt32(strValor.Substring(3, 3)) > 1)
                            valor_por_extenso += " BILHÕES" + ((Convert.ToDecimal(strValor.Substring(6, 9)) > 0) ? " E " : string.Empty);
                    }
                    else if (i == 6 & valor_por_extenso != string.Empty)
                    {
                        if (Convert.ToInt32(strValor.Substring(6, 3)) == 1)
                            valor_por_extenso += " MILHÃO" + ((Convert.ToDecimal(strValor.Substring(9, 6)) > 0) ? " E " : string.Empty);
                        else if (Convert.ToInt32(strValor.Substring(6, 3)) > 1)
                            valor_por_extenso += " MILHÕES" + ((Convert.ToDecimal(strValor.Substring(9, 6)) > 0) ? " E " : string.Empty);
                    }
                    else if (i == 9 & valor_por_extenso != string.Empty)
                        if (Convert.ToInt32(strValor.Substring(9, 3)) > 0)
                            valor_por_extenso += " MIL" + ((Convert.ToDecimal(strValor.Substring(12, 3)) > 0) ? " E " : string.Empty);

                    if (i == 12)
                    {
                        if (valor_por_extenso.Length > 8)
                            if (valor_por_extenso.Substring(valor_por_extenso.Length - 6, 6) == "BILHÃO" | valor_por_extenso.Substring(valor_por_extenso.Length - 6, 6) == "MILHÃO")
                                valor_por_extenso += " DE";
                            else
                                if (valor_por_extenso.Substring(valor_por_extenso.Length - 7, 7) == "BILHÕES" | valor_por_extenso.Substring(valor_por_extenso.Length - 7, 7) == "MILHÕES" | valor_por_extenso.Substring(valor_por_extenso.Length - 8, 7) == "TRILHÕES")
                                valor_por_extenso += " DE";
                            else
                                    if (valor_por_extenso.Substring(valor_por_extenso.Length - 8, 8) == "TRILHÕES")
                                valor_por_extenso += " DE";

                        if (Convert.ToInt64(strValor.Substring(0, 15)) == 1)
                            valor_por_extenso += " REAL";
                        else if (Convert.ToInt64(strValor.Substring(0, 15)) > 1)
                            valor_por_extenso += " REAIS";

                        if (Convert.ToInt32(strValor.Substring(16, 2)) > 0 && valor_por_extenso != string.Empty)
                            valor_por_extenso += " E ";
                    }

                    if (i == 15)
                        if (Convert.ToInt32(strValor.Substring(16, 2)) == 1)
                            valor_por_extenso += " CENTAVO";
                        else if (Convert.ToInt32(strValor.Substring(16, 2)) > 1)
                            valor_por_extenso += " CENTAVOS";
                }
                return valor_por_extenso;
            }
        }

        static string escreva_parte(decimal valor)
        {
            if (valor <= 0)
                return string.Empty;
            else
            {
                string montagem = string.Empty;
                if (valor > 0 & valor < 1)
                {
                    valor *= 100;
                }
                string strValor = valor.ToString("000");
                int a = Convert.ToInt32(strValor.Substring(0, 1));
                int b = Convert.ToInt32(strValor.Substring(1, 1));
                int c = Convert.ToInt32(strValor.Substring(2, 1));

                if (a == 1) montagem += (b + c == 0) ? "CEM" : "CENTO";
                else if (a == 2) montagem += "DUZENTOS";
                else if (a == 3) montagem += "TREZENTOS";
                else if (a == 4) montagem += "QUATROCENTOS";
                else if (a == 5) montagem += "QUINHENTOS";
                else if (a == 6) montagem += "SEISCENTOS";
                else if (a == 7) montagem += "SETECENTOS";
                else if (a == 8) montagem += "OITOCENTOS";
                else if (a == 9) montagem += "NOVECENTOS";

                if (b == 1)
                {
                    if (c == 0) montagem += ((a > 0) ? " E " : string.Empty) + "DEZ";
                    else if (c == 1) montagem += ((a > 0) ? " E " : string.Empty) + "ONZE";
                    else if (c == 2) montagem += ((a > 0) ? " E " : string.Empty) + "DOZE";
                    else if (c == 3) montagem += ((a > 0) ? " E " : string.Empty) + "TREZE";
                    else if (c == 4) montagem += ((a > 0) ? " E " : string.Empty) + "QUATORZE";
                    else if (c == 5) montagem += ((a > 0) ? " E " : string.Empty) + "QUINZE";
                    else if (c == 6) montagem += ((a > 0) ? " E " : string.Empty) + "DEZESSEIS";
                    else if (c == 7) montagem += ((a > 0) ? " E " : string.Empty) + "DEZESSETE";
                    else if (c == 8) montagem += ((a > 0) ? " E " : string.Empty) + "DEZOITO";
                    else if (c == 9) montagem += ((a > 0) ? " E " : string.Empty) + "DEZENOVE";
                }
                else if (b == 2) montagem += ((a > 0) ? " E " : string.Empty) + "VINTE";
                else if (b == 3) montagem += ((a > 0) ? " E " : string.Empty) + "TRINTA";
                else if (b == 4) montagem += ((a > 0) ? " E " : string.Empty) + "QUARENTA";
                else if (b == 5) montagem += ((a > 0) ? " E " : string.Empty) + "CINQUENTA";
                else if (b == 6) montagem += ((a > 0) ? " E " : string.Empty) + "SESSENTA";
                else if (b == 7) montagem += ((a > 0) ? " E " : string.Empty) + "SETENTA";
                else if (b == 8) montagem += ((a > 0) ? " E " : string.Empty) + "OITENTA";
                else if (b == 9) montagem += ((a > 0) ? " E " : string.Empty) + "NOVENTA";

                if (strValor.Substring(1, 1) != "1" & c != 0 & montagem != string.Empty) montagem += " E ";

                if (strValor.Substring(1, 1) != "1")
                    if (c == 1) montagem += "UM";
                    else if (c == 2) montagem += "DOIS";
                    else if (c == 3) montagem += "TRÊS";
                    else if (c == 4) montagem += "QUATRO";
                    else if (c == 5) montagem += "CINCO";
                    else if (c == 6) montagem += "SEIS";
                    else if (c == 7) montagem += "SETE";
                    else if (c == 8) montagem += "OITO";
                    else if (c == 9) montagem += "NOVE";

                return montagem;
            }
        }
    }
}
