using System;

namespace Vovin.CmcLibNet.Extensions
{
    /// <summary>
    /// String extension methods
    /// </summary>
    internal static class StringExtensions
    {
        /// <summary>
        /// Returns number of occurrences of a character in a string
        /// </summary>
        /// <param name="str">String to search in.</param>
        /// <param name="c">Character to count.</param>
        /// <returns></returns>
        internal static int CountChar(this string str, char c)
        {
            int counter = 0;
            foreach (char x in str)
            {
                if (x == c) { counter++; }
            }
            return counter;
        }

        /// <summary>
        /// Returns Nth index of character in a string
        /// </summary>
        /// <param name="input">String to search in</param>
        /// <param name="value">String to search for</param>
        /// <param name="startIndex">Start index</param>
        /// <param name="nth">Nth occurrence</param>
        /// <returns></returns>
        /// <exception cref="NotSupportedException"></exception>
        internal static int IndexOfNthChar(this string input, char value, int startIndex, int nth)
        {
            if (nth < 1)
                throw new NotSupportedException("Param 'nth' must be greater than 0!");
            if (nth == 1)
                return input.IndexOf(value, startIndex);

            return input.IndexOfNthChar(value, input.IndexOf(value, startIndex) + 1, --nth);
        }

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
    }
}
