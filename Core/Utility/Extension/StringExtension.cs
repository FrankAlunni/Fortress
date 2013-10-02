// --------------------------------------------------------------------------------------------------------------------
// <copyright file="StringExtension.cs" company="B1C Canada Inc.">
//   Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <summary>
//   Extension to the string class
// </summary>
// --------------------------------------------------------------------------------------------------------------------
using System.IO;
using System.Xml;

namespace B1C.Utility.Extension
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using System.Text.RegularExpressions;

    /// <summary>
    /// Extension to the string class
    /// </summary>
    public static class StringExtension
    {
        /// <summary>
        /// Converts the XML to formatted XML.
        /// </summary>
        /// <param name="input">The input.</param>
        /// <returns>The Formatted XML</returns>
        public static string ToXml(this string input)
        {
            // load unformatted xml into a dom
            var xd = new XmlDocument();
            xd.LoadXml(input);

            // will hold formatted xml
            var sb = new StringBuilder();

            // pumps the formatted xml into the StringBuilder above
            var sw = new StringWriter(sb);

            // does the formatting
            XmlTextWriter xtw = null;

            try
            {
                // point the xtw at the StringWriter
                xtw = new XmlTextWriter(sw) { Formatting = Formatting.Indented };

                // get the dom to dump its contents into the xtw 
                xd.WriteTo(xtw);
            }
            finally
            {
                // clean up even if error
                if (xtw != null)
                {
                    xtw.Close();
                }
            }

            // return the formatted xml
            return sb.ToString();
        }

        /// <summary>
        /// Convert the input string into a comma separated list of words.
        /// </summary>
        /// <param name="input">The strign to analize</param>
        /// <param name="wordDelimiter">The word delimiter.</param>
        /// <returns>A comma separated list of words</returns>
        public static string ToCsvWords(this string input, char wordDelimiter)
        {
            var words = input.ToWords();

            string returnValue = string.Empty;

            foreach (string word in words)
            {
                if (string.IsNullOrEmpty(returnValue))
                {
                    returnValue = wordDelimiter + word + wordDelimiter;
                }
                else
                {
                    returnValue += "," + wordDelimiter + word + wordDelimiter;
                }
            }

            return returnValue;
        }

        /// <summary>
        /// Creates a list of words from the input string
        /// </summary>
        /// <param name="input">The string to examine</param>
        /// <returns>
        /// Returns a list of words from the input string
        /// </returns>
        public static List<string> ToWords(this string input)
        {
            // char[] delimiters = { ',', '.', '"', '-', ' ', ':', ';', '\n', '!', '?', '\r', '\t' };
            char[] delimiters = { ',', '.', '"', ' ', ':', ';', '\n', '!', '?', '\r', '\t' };

            Array arrayWords = input.Split(delimiters);            
            var words = new List<string>();
            
            foreach (string word in arrayWords)
            {
                if (word.Trim().Length > 0)
                {
                    words.Add(word.Trim());
                }
            }

            return words;
        }

        /// <summary>
        /// Base64 encodes a string.
        /// </summary>
        /// <param name="input">The string to examine</param>
        /// <returns>A base64 encoded string</returns>
        public static string Base64StringEncode(this string input)
        {
            byte[] encbuff = Encoding.UTF8.GetBytes(input);
            return Convert.ToBase64String(encbuff);
        }

        /// <summary>
        /// Base64 decodes a string.
        /// </summary>
        /// <param name="input">A base64 encoded string</param>
        /// <returns>A decoded string</returns>
        public static string Base64StringDecode(this string input)
        {
            byte[] decbuff = Convert.FromBase64String(input);
            return Encoding.UTF8.GetString(decbuff);
        }

        /// <summary>
        /// A case insenstive replace function.
        /// </summary>
        /// <param name="input">The string to examine.</param>
        /// <param name="newValue">The value to replace.</param>
        /// <param name="oldValue">The new value to be inserted</param>
        /// <returns>A string replaced</returns>
        public static string CaseInsenstiveReplace(this string input, string newValue, string oldValue)
        {
            var regEx = new Regex(oldValue, RegexOptions.IgnoreCase | RegexOptions.Multiline);
            return regEx.Replace(input, newValue);
        }

        /// <summary>
        /// Replaces the first occurence of a string with the replacement value. The Replace
        /// is case senstive
        /// </summary>
        /// <param name="input">The string to examine</param>
        /// <param name="oldValue">The value to replace</param>
        /// <param name="newValue">the new value to be inserted</param>
        /// <returns>The string replaced</returns>
        public static string ReplaceFirst(this string input, string oldValue, string newValue)
        {
            var regEx = new Regex(oldValue, RegexOptions.Multiline);
            return regEx.Replace(input, newValue, 1);
        }

        /// <summary>
        /// Replaces the last occurence of a string with the replacement value.
        /// The replace is case senstive.
        /// </summary>
        /// <param name="input">The string to examine</param>
        /// <param name="oldValue">The value to replace</param>
        /// <param name="newValue">the new value to be inserted</param>
        /// <returns>The string replaced</returns>
        public static string ReplaceLast(this string input, string oldValue, string newValue)
        {
            int index = input.LastIndexOf(oldValue);
            if (index < 0)
            {
                return input;
            }

            var sb = new StringBuilder(input.Length - oldValue.Length + newValue.Length);
            sb.Append(input.Substring(0, index));
            sb.Append(newValue);
            sb.Append(input.Substring(index + oldValue.Length, input.Length - index - oldValue.Length));

            return sb.ToString();
        }

        /// <summary>
        /// Removes all the words passed in the filter words 
        /// parameters. The replace is NOT case
        /// sensitive.
        /// </summary>
        /// <param name="input">The string to search.</param>
        /// <param name="filterWords">The words to 
        /// repace in the input string.</param>
        /// <returns>A filtered string.</returns>
        public static string FilterWords(this string input, params string[] filterWords)
        {
            return input.FilterWords(char.MinValue, filterWords);
        }

        /// <summary>
        /// Removes all the words passed in the filter words
        /// parameters. The replace is NOT case
        /// sensitive.
        /// </summary>
        /// <param name="input">The string to examine</param>
        /// <param name="mask">The mask to use</param>
        /// <param name="filterWords">The filter words</param>
        /// <returns>A filtered string.</returns>
        public static string FilterWords(this string input, char mask, params string[] filterWords)
        {
            string stringMask = mask == char.MinValue ?
               string.Empty : mask.ToString();
            string totalMask = stringMask;

            foreach (string s in filterWords)
            {
                var regEx = new Regex(s, RegexOptions.IgnoreCase | RegexOptions.Multiline);

                if (stringMask.Length > 0)
                {
                    for (int i = 1; i < s.Length; i++)
                    {
                        totalMask += stringMask;
                    }
                }

                input = regEx.Replace(input, totalMask);

                totalMask = stringMask;
            }

            return input;
        }

        /// <summary>
        /// Left pads the passed input using the passed pad string
        /// for the total number of spaces.  It will not
        /// cut-off the pad even if it
        /// causes the string to exceed the total width.
        /// </summary>
        /// <param name="input">The string to pad.</param>
        /// <param name="pad">The string to uses as padding.</param>
        /// <param name="totalWidth">The total width.</param>
        /// <returns>A padded string.</returns>
        public static string PadLeft(this string input, string pad, int totalWidth)
        {
            return input.PadLeft(pad, totalWidth, false);
        }

        /// <summary>
        /// Left pads the passed input using the passed pad string
        /// for the total number of spaces.  It will
        /// cut-off the pad so that
        /// the string does not exceed the total width.
        /// </summary>
        /// <param name="input">The string to pad.</param>
        /// <param name="pad">The string to uses as padding.</param>
        /// <param name="totalWidth">The total width.</param>
        /// <param name="cutOff">if set to <c>true</c> [cut off].</param>
        /// <returns>A padded string.</returns>
        public static string PadLeft(this string input, string pad, int totalWidth, bool cutOff)
        {
            if (input.Length >= totalWidth)
            {
                return input;
            }

            string paddedString = input;

            while (paddedString.Length < totalWidth)
            {
                paddedString += pad;
            }

            // trim the excess.
            if (cutOff)
            {
                paddedString = paddedString.Substring(0, totalWidth);
            }

            return paddedString;
        }

        /// <summary>
        /// Right pads the passed input using the passed pad string
        /// for the total number of spaces.  It will not 
        /// cut-off the pad even if it 
        /// causes the string to exceed the total width.
        /// </summary>
        /// <param name="input">The string to pad.</param>
        /// <param name="pad">The string to uses as padding.</param>
        /// <param name="totalWidth">The total number to pad the string.</param>
        /// <returns>A padded string.</returns>
        public static string PadRight(this string input, string pad, int totalWidth)
        {
            return input.PadRight(pad, totalWidth, false);
        }

        /// <summary>
        /// Right pads the passed input using the passed pad string
        /// for the total number of spaces.  It will cut-off
        /// the pad so that
        /// the string does not exceed the total width.
        /// </summary>
        /// <param name="input">The string to pad.</param>
        /// <param name="pad">The string to uses as padding.</param>
        /// <param name="totalWidth">The total number to  pad the string.</param>
        /// <param name="cutOff">if set to <c>true</c> [cut off].</param>
        /// <returns>A padded string.</returns>
        public static string PadRight(this string input, string pad, int totalWidth, bool cutOff)
        {
            if (input.Length >= totalWidth)
            {
                return input;
            }

            string paddedString = string.Empty;

            while (paddedString.Length < totalWidth - input.Length)
            {
                paddedString += pad;
            }

            // trim the excess.
            if (cutOff)
            {
                paddedString = paddedString.Substring(0, totalWidth - input.Length);
            }

            paddedString += input;

            return paddedString;
        }

        /// <summary>
        /// Removes the new line (\n) and carriage return (\r) symbols.
        /// </summary>
        /// <param name="input">The string to search.</param>
        /// <returns>A string without New Lines</returns>
        public static string RemoveNewLines(this string input)
        {
            return input.RemoveNewLines(false);
        }

        /// <summary>
        /// Removes the new line (\n) and carriage return 
        /// (\r) symbols.
        /// </summary>
        /// <param name="input">The string to search.</param>
        /// <param name="addSpace">If true, adds a space 
        /// (" ") for each newline and carriage
        /// return found.</param>
        /// <returns>A string without New Lines</returns>
        public static string RemoveNewLines(this string input, bool addSpace)
        {
            string replace = string.Empty;
            if (addSpace)
            {
                replace = " ";
            }

            const string Pattern = @"[\r|\n]";
            var regEx = new Regex(Pattern, RegexOptions.IgnoreCase | RegexOptions.Multiline);

            return regEx.Replace(input, replace);
        }

        /// <summary>
        /// Reverse a string.
        /// </summary>
        /// <param name="input">The string to reverse</param>
        /// <returns>A string to examine</returns>
        /// <remarks>Thanks to  Alois Kraus for pointing out an issue
        /// with an earlier version of this method and to Justin Roger's 
        /// site (http://weblogs.asp.net/justin_rogers/archive/2004/06/10/153175.aspx)
        /// for helping me to improve that previous method.</remarks>
        public static string Reverse(this string input)
        {
            var reverse = new char[input.Length];
            for (int i = 0, k = input.Length - 1; i < input.Length; i++, k--)
            {
                if (char.IsSurrogate(input[k]))
                {
                    reverse[i + 1] = input[k--];
                    reverse[i++] = input[k];
                }
                else
                {
                    reverse[i] = input[k];
                }
            }

            return new string(reverse);
        }

        /// <summary>
        /// Converts a string to sentence case.
        /// </summary>
        /// <param name="input">The string to convert.</param>
        /// <returns>A sentenced cased string</returns>
        public static string SentenceCase(this string input)
        {
            if (input.Length < 1)
            {
                return input;
            }

            string sentence = input.ToLower();
            return sentence[0].ToString().ToUpper() + sentence.Substring(1);
        }

        /// <summary>
        /// Converts a string to title case.
        /// </summary>
        /// <param name="input">The string to convert.</param>
        /// <returns>A titled cased string</returns>
        public static string TitleCase(this string input)
        {
            return input.TitleCase(true);
        }

        /// <summary>
        /// Converts a string to title case.
        /// </summary>
        /// <param name="input">The string to convert.</param>
        /// <param name="ignoreShortWords">If true, 
        /// does not capitalize words like
        /// "a", "is", "the", etc.</param>
        /// <returns>A titled cased string</returns>
        public static string TitleCase(this string input, bool ignoreShortWords)
        {
            List<string> ignoreWords = null;
            if (ignoreShortWords)
            {
                // Add more ignore words if necessary
                ignoreWords = new List<string> { "a", "is", "was", "the" };
            }

            string[] tokens = input.Split(' ');
            var sb = new StringBuilder(input.Length);
            foreach (string s in tokens)
            {
                if (ignoreShortWords && s != tokens[0] && ignoreWords.Contains(s.ToLower()))
                {
                    sb.Append(s + " ");
                }
                else
                {
                    sb.Append(s[0].ToString().ToUpper());
                    sb.Append(s.Substring(1).ToLower());
                    sb.Append(" ");
                }
            }

            return sb.ToString().Trim();
        }

        /// <summary>
        /// Removes multiple spaces between words
        /// </summary>
        /// <param name="input">The string to trim.</param>
        /// <returns>A string without multiple spaces between words.</returns>
        public static string TrimIntraWords(this string input)
        {
            var regEx = new Regex(@"[\s]+");
            return regEx.Replace(input, " ");
        }

        /// <summary>
        /// Wraps the passed string at the 
        /// at the next whitespace on or after the 
        /// total charCount has been reached
        /// for that line.  Uses the environment new line
        /// symbol for the break text.
        /// </summary>
        /// <param name="input">The string to wrap.</param>
        /// <param name="charCount">The number of characters 
        /// per line.</param>
        /// <returns>A string word wrapped.</returns>
        public static string WordWrap(this string input, int charCount)
        {
            return input.WordWrap(charCount, false, Environment.NewLine);
        }

        /// <summary>
        /// Wraps the passed string at the total 
        /// number of characters (if cuttOff is true)
        /// or at the next whitespace (if cutOff is false).
        /// Uses the environment new line
        /// symbol for the break text.
        /// </summary>
        /// <param name="input">The string to wrap.</param>
        /// <param name="charCount">The number of characters 
        /// per line.</param>
        /// <param name="cutOff">If true, will break in 
        /// the middle of a word.</param>
        /// <returns>A string word wrapped.</returns>
        public static string WordWrap(this string input, int charCount, bool cutOff)
        {
            return input.WordWrap(charCount, cutOff, Environment.NewLine);
        }

        /// <summary>
        /// Wraps the passed string at the total number 
        /// of characters (if cuttOff is true)
        /// or at the next whitespace (if cutOff is false).
        /// Uses the passed breakText
        /// for lineBreaks.
        /// </summary>
        /// <param name="input">The string to wrap.</param>
        /// <param name="charCount">The number of 
        /// characters per line.</param>
        /// <param name="cutOff">If true, will break in 
        /// the middle of a word.</param>
        /// <param name="breakText">The line break text to use.</param>
        /// <returns>A string word wrapped.</returns>
        public static string WordWrap(this string input, int charCount, bool cutOff, string breakText)
        {
            var sb = new StringBuilder(input.Length + 100);
            int counter = 0;

            if (cutOff)
            {
                while (counter < input.Length)
                {
                    if (input.Length > counter + charCount)
                    {
                        sb.Append(input.Substring(counter, charCount));
                        sb.Append(breakText);
                    }
                    else
                    {
                        sb.Append(input.Substring(counter));
                    }

                    counter += charCount;
                }
            }
            else
            {
                string[] strings = input.Split(' ');
                for (int i = 0; i < strings.Length; i++)
                {
                    // added one to represent the space.
                    counter += strings[i].Length + 1;
                    if (i != 0 && counter > charCount)
                    {
                        sb.Append(breakText);
                        counter = 0;
                    }

                    sb.Append(strings[i] + ' ');
                }
            }

            // to get rid of the extra space at the end.
            return sb.ToString().TrimEnd();
        }

        /// <summary>
        /// Determines whether the specified input is email.
        /// </summary>
        /// <param name="input">The input.</param>
        /// <returns>
        /// <c>true</c> if the specified input is email; otherwise, <c>false</c>.
        /// </returns>
        public static bool IsEmail(this string input)
        {
            const string Pattern = @"^([0-9a-zA-Z]+[-._+&])*[0-9a-zA-Z]+@([-0-9a-zA-Z]+[.])+[a-zA-Z]{2,6}$";
            return Regex.Match(input, Pattern).Success;
        }

        /// <summary>
        /// Determines whether [is us zip code] [the specified input].
        /// </summary>
        /// <param name="input">The input.</param>
        /// <returns>
        /// <c>true</c> if [is us zip code] [the specified input]; otherwise, <c>false</c>.
        /// </returns>
        public static bool IsUsZipCode(this string input)
        {
            const string Pattern = @"^\d{5}$";
            return Regex.Match(input, Pattern).Success;
        }

        /// <summary>
        /// Determines whether [is us zip code plus4 mandatory] [the specified input].
        /// </summary>
        /// <param name="input">The input.</param>
        /// <returns>
        /// <c>true</c> if [is us zip code plus4 mandatory] [the specified input]; otherwise, <c>false</c>.
        /// </returns>
        public static bool IsUsZipCodePlus4Mandatory(this string input)
        {
            const string Pattern = @"^\d{5}((-|\s)?\d{4})$";
            return Regex.Match(input, Pattern).Success;
        }

        /// <summary>
        /// Determines whether [is us zip code plus4 optional] [the specified input].
        /// </summary>
        /// <param name="input">The input.</param>
        /// <returns>
        /// <c>true</c> if [is us zip code plus4 optional] [the specified input]; otherwise, <c>false</c>.
        /// </returns>
        public static bool IsUsZipCodePlus4Optional(this string input)
        {
            const string Pattern = @"^\d{5}((-|\s)?\d{4})?$";
            return Regex.Match(input, Pattern).Success;
        }

        /// <summary>
        /// Determines whether [is canada postal code] [the specified input].
        /// </summary>
        /// <example>
        /// Allows: A1A 1A1 or A1A1A1
        /// </example>
        /// <param name="input">The input.</param>
        /// <returns>
        /// <c>true</c> if [is canada postal code] [the specified input]; otherwise, <c>false</c>.
        /// </returns>
        public static bool IsCanadaPostalCode(this string input)
        {
            const string Pattern = @"^([ABCEGHJKLMNPRSTVXY]\d[ABCEGHJKLMNPRSTVWXYZ])\ {0,1}(\d[ABCEGHJKLMNPRSTVWXYZ]\d)$";
            return Regex.Match(input, Pattern).Success;
        }

        /// <summary>
        /// Determines whether [is phone number] [the specified input].
        /// </summary>
        /// <example>
        /// Allows: 324-234-3433, 3242343434, (234)234-234, (234) 234-2343
        /// </example>
        /// <param name="input">The input.</param>
        /// <returns>
        /// <c>true</c> if [is phone number] [the specified input]; otherwise, <c>false</c>.
        /// </returns>
        public static bool IsPhoneNumber(this string input)
        {
            const string Pattern = @"^([\(]{1}[0-9]{3}[\)]{1}[\.| |\-]{0,1}|^[0-9]{3}[\.|\-| ]?)?[0-9]{3}(\.|\-| )?[0-9]{4}$";
            return Regex.Match(input, Pattern).Success;
        }

        /// <summary>
        /// Determines whether the specified input is URL.
        /// </summary>
        /// <example>
        /// Allows: http://www.yahoo.com, https://www.yahoo.com, ftp://www.yahoo.com
        /// </example>
        /// <param name="input">The input.</param>
        /// <returns>
        /// <c>true</c> if the specified input is URL; otherwise, <c>false</c>.
        /// </returns>
        public static bool IsUrl(this string input)
        {
            const string Pattern = @"^(?<Protocol>\w+):\/\/(?<Domain>[\w.]+\/?)\S*$";
            return Regex.Match(input, Pattern).Success;
        }

        /// <summary>
        /// Determines whether [is ip address] [the specified input].
        /// </summary>
        /// <example>
        /// Allows: 123.123.123.123, 192.168.1.1
        /// </example>
        /// <param name="input">The input.</param>
        /// <returns>
        /// <c>true</c> if [is ip address] [the specified input]; otherwise, <c>false</c>.
        /// </returns>
        public static bool IsIpAddress(this string input)
        {
            const string Pattern = @"^(?<First>2[0-4]\d|25[0-5]|[01]?\d\d?)\.(?<Second>2[0-4]\d|25[0-5]|[01]?\d\d?)\.(?<Third>2[0-4]\d|25[0-5]|[01]?\d\d?)\.(?<Fourth>2[0-4]\d|25[0-5]|[01]?\d\d?)$";
            return Regex.Match(input, Pattern).Success;
        }

        /// <summary>
        /// This matches a date in the format mm/dd/yy
        /// </summary>
        /// <example>
        /// Allows: 01/05/05, 12/30/99, 04/11/05
        /// Does not allow: 01/05/2000, 2/2/02
        /// </example> 
        /// <param name="input">The input.</param>
        /// <returns>
        /// <c>true</c> if [is date mm dd yy] [the specified input]; otherwise, <c>false</c>.
        /// </returns>
        public static bool IsDateMmDdYy(this string input)
        {
            const string Pattern = @"^(1[0-2]|0[1-9])/(([1-2][0-9]|3[0-1]|0[1-9])/\d\d)$"; 
            return Regex.Match(input, Pattern).Success;
        }

        /// <summary>
        /// This matches a date in the format mm/yy
        /// </summary>
        /// <example>
        /// Allows: 01/05, 11/05, 04/99
        /// Does not allow: 1/05, 13/05, 00/05
        /// </example>
        /// <param name="input">The input.</param>
        /// <returns>
        /// <c>true</c> if [is date mm yy] [the specified input]; otherwise, <c>false</c>.
        /// </returns>
        public static bool IsDateMmYy(this string input)
        {
            const string Pattern = @"^((0[1-9])|(1[0-2]))\/(\d{2})$";
            return Regex.Match(input, Pattern).Success;
        }

        /// <summary>
        ///     This matches only numbers(no decimals)
        /// </summary>
        /// <example>
        /// Allows: 0, 1, 123, 4232323, 1212322
        /// </example>
        /// <param name="input">The input.</param>
        /// <returns>
        /// <c>true</c> if [is number only] [the specified input]; otherwise, <c>false</c>.
        /// </returns>
        public static bool IsNumberOnly(this string input)
        {
            const string Pattern = @"^((0[1-9])|(1[0-2]))\/(\d{2})$";
            return Regex.Match(input, Pattern).Success;
        }

        /// <summary>
        /// This matches any string with only alpha characters upper or lower case(A-Z)
        /// </summary>
        /// <example>
        /// Allows: abc, ABC, abCd, AbCd
        /// Does not allow: A C, abc!, (a,b)
        /// </example>
        /// <param name="input">The input.</param>
        /// <returns>
        /// <c>true</c> if [is alpha only] [the specified input]; otherwise, <c>false</c>.
        /// </returns>
        public static bool IsAlphaOnly(this string input)
        {
            const string Pattern = @"^[a-zA-Z]+$";
            return Regex.Match(input, Pattern).Success;
        }

        /// <summary>
        /// Determines whether [is upper case] [the specified input].
        /// </summary>
        /// <param name="input">The input.</param>
        /// <returns>
        /// <c>true</c> if [is upper case] [the specified input]; otherwise, <c>false</c>.
        /// </returns>
        public static bool IsUpperCase(this string input)
        {
            const string Pattern = @"^[A-Z]+$";
            return Regex.Match(input, Pattern).Success;
        }

        /// <summary>
        /// Determines whether [is lower case] [the specified input].
        /// </summary>
        /// <param name="input">The input.</param>
        /// <returns>
        /// <c>true</c> if [is lower case] [the specified input]; otherwise, <c>false</c>.
        /// </returns>
        public static bool IsLowerCase(this string input)
        {
            const string Pattern = @"^[a-z]+$";
            return Regex.Match(input, Pattern).Success;
        }

        /// <summary>
        /// Ensures the string only contains alpha-numeric characters, and
        /// not punctuation, spaces, line breaks, etc.
        /// </summary>
        /// <example>
        /// Allows: ab2c, 112ABC, ab23Cd
        /// Does not allow: A C, abc!, a.a
        /// </example>
        /// <param name="input">The input.</param>
        /// <returns>
        /// <c>true</c> if [is alpha numeric] [the specified input]; otherwise, <c>false</c>.
        /// </returns>
        public static bool IsAlphaNumeric(this string input)
        {
            const string Pattern = @"^[a-zA-Z0-9]+$";
            return Regex.Match(input, Pattern).Success;
        }

        /// <summary>
        /// Validates US Currency.  Requires $ sign
        /// Allows for optional commas and decimal. 
        /// No leading zeros. 
        /// </summary>
        /// <example>Allows: $100,000 or $10000.00 or $10.00 or $.10 or $0 or $0.00
        /// Does not allow: $0.10 or 10.00 or 10,000
        /// </example>
        /// <param name="input">The input.</param>
        /// <returns>
        /// <c>true</c> if the specified input is currency; otherwise, <c>false</c>.
        /// </returns>
        public static bool IsCurrency(this string input)
        {
            const string Pattern = @"^\$(([1-9]\d*|([1-9]\d{0,2}(\,\d{3})*))(\.\d{1,2})?|(\.\d{1,2}))$|^\$[0](.00)?$";
            return Regex.Match(input, Pattern).Success;
        }

        /// <summary>
        /// Matches major credit cards including: Visa (length 16, prefix 4); 
        /// Mastercard (length 16, prefix 51-55);
        /// Diners Club/Carte Blanche (length 14, prefix 36, 38, or 300-305); 
        /// Discover (length 16, prefix 6011); 
        /// American Express (length 15, prefix 34 or 37). 
        /// Saves the card type as a named group to facilitate further validation 
        /// against a "card type" checkbox in a program. 
        /// All 16 digit formats are grouped 4-4-4-4 with an optional hyphen or space 
        /// between each group of 4 digits. 
        /// The American Express format is grouped 4-6-5 with an optional hyphen or space 
        /// between each group of digits. 
        /// Formatting characters must be consistant, i.e. if two groups are separated by a hyphen, 
        /// all groups must be separated by a hyphen for a match to occur.
        /// </summary>
        /// <param name="input">The input.</param>
        /// <returns>
        /// <c>true</c> if [is credit card] [the specified input]; otherwise, <c>false</c>.
        /// </returns>
        public static bool IsCreditCard(this string input)
        {
            const string Pattern = @"^(?:(?<Visa>4\d{3})|(?<Mastercard>5[1-5]\d{2})|(?<Discover>6011)|(?<DinersClub>(?:3[68]\d{2})|(?:30[0-5]\d))|(?<AmericanExpress>3[47]\d{2}))([ -]?)(?(DinersClub)(?:\d{6}\1\d{4})|(?(AmericanExpress)(?:\d{6}\1\d{5})|(?:\d{4}\1\d{4}\1\d{4})))$";
            return Regex.Match(input, Pattern).Success;
        }

        /// <summary>
        /// Matches social security in the following format xxx-xx-xxxx
        /// where x is a number
        /// </summary>
        /// <example>
        /// Allows: 123-45-6789, 232-432-1212
        /// </example>
        /// <param name="input">The input.</param>
        /// <returns>
        /// <c>true</c> if [is social security number] [the specified input]; otherwise, <c>false</c>.
        /// </returns>
        public static bool IsSocialSecurityNumber(this string input)
        {
            const string Pattern = @"^\d{3}-\d{2}-\d{4}$";
            return Regex.Match(input, Pattern).Success;
        }

        /// <summary>
        /// Extracts the between XML tag.
        /// </summary>
        /// <param name="input">The input.</param>
        /// <param name="xmlTag">The XML tag.</param>
        /// <returns>
        /// The extracted string between the XML tags.
        /// </returns>
        public static string ExtractBetweenXmlTag(this string input, string xmlTag)
        {
            var startTag = "<" + xmlTag + ">";
            var endTag = "</" + xmlTag + ">";
            string result = ExtractBetweenTags(input, startTag, endTag);
            if (string.IsNullOrEmpty(result))
            {
                endTag = "/>";
                result = ExtractBetweenTags(input, startTag, endTag);
            }

            return result;
        }

        /// <summary>
        /// Extracts the between tags.
        /// </summary>
        /// <param name="input">The input.</param>
        /// <param name="startTag">The start tag.</param>
        /// <param name="endTag">The end tag.</param>
        /// <returns>The extracted string.</returns>
        public static string ExtractBetweenTags(this string input, string startTag, string endTag)
        {
            try
            {
                int startIndex = input.IndexOf(startTag) + startTag.Length;
                int endIndex = input.IndexOf(endTag, startIndex);
                return input.Substring(startIndex, endIndex - startIndex);
            }
            catch
            {
                return string.Empty;
            }
        } 
    }
}
