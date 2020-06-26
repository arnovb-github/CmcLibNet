using System;

namespace Vovin.CmcLibNet.Extensions
{
    /// <summary>
    /// String extension methods
    /// </summary>
    internal static class StringExtensions
    {
        /// <summary>
        /// Surround string with character.
        /// </summary>
        /// <param name="input"><c>string</c> to enclose.</param>
        /// <param name="encloseWith"><c>char</c> to enclose with.</param>
        /// <param name="number">Number of chars to enclose with (up to 255).</param>
        /// <returns><c>string</c> enclosed by char(s).</returns>
        internal static string EncloseWithChar(this string input, char encloseWith, byte number = 1)
        {
            string retval = input;
            for (int i = 0; i < number; i++)
            {
                retval = encloseWith + retval + encloseWith;
            }
            return retval;
        }

        internal static string Left(this string str, int length)
        {
            return str.Substring(0, Math.Min(length, str.Length));
        }

        internal static string Right(this string str, int length)
        {
            return str.Substring(str.Length - Math.Min(length, str.Length));
        }

        internal static int CountChar(this string str, char character)
        {
            int count = 0;
            for (int i = 0; i < str?.Length; i++) 
            {
                if (str[i].Equals(character))
                    count++;
            }
            return count;
        }
    }
}
