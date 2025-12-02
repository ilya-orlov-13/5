using System;
using System.Collections.Generic;
using System.Linq;

public static class ConsoleValidator
{
    #region Char Validation

    [Flags]
    public enum CharValidationOptions
    {
        None = 0,
        IgnoreCase = 1,
        LettersOnly = 2,
        DigitsOnly = 4,
        LettersOrDigits = 8,
        NoWhitespace = 16
    }

    public static char ReadValidatedChar(string prompt, CharValidationOptions options = CharValidationOptions.None,
        params char[] allowedChars)
    {
        while (true)
        {
            try
            {
                Console.Write(prompt);
                string input = Console.ReadLine();

                if (string.IsNullOrEmpty(input))
                {
                    Console.WriteLine("Ошибка: Ввод не может быть пустым. Попробуйте снова.");
                    continue;
                }

                if (input.Length != 1)
                {
                    Console.WriteLine($"Ошибка: Нужно ввести ровно один символ. Вы ввели {input.Length}. Попробуйте снова.");
                    continue;
                }

                char result = input[0];

                if (options.HasFlag(CharValidationOptions.IgnoreCase))
                {
                    result = char.ToUpperInvariant(result);
                }

                if (options.HasFlag(CharValidationOptions.LettersOnly) && !char.IsLetter(result))
                {
                    Console.WriteLine("Ошибка: Требуется буква. Попробуйте снова.");
                    continue;
                }

                if (options.HasFlag(CharValidationOptions.DigitsOnly) && !char.IsDigit(result))
                {
                    Console.WriteLine("Ошибка: Требуется цифра. Попробуйте снова.");
                    continue;
                }

                if (options.HasFlag(CharValidationOptions.LettersOrDigits) && !char.IsLetterOrDigit(result))
                {
                    Console.WriteLine("Ошибка: Требуется буква или цифра. Попробуйте снова.");
                    continue;
                }

                if (options.HasFlag(CharValidationOptions.NoWhitespace) && char.IsWhiteSpace(result))
                {
                    Console.WriteLine("Ошибка: Запрещены пробельные символы. Попробуйте снова.");
                    continue;
                }

                if (allowedChars.Length > 0 && !allowedChars.Contains(result))
                {
                    Console.WriteLine($"Ошибка: Символ должен быть одним из: {string.Join(", ", allowedChars)}. Попробуйте снова.");
                    continue;
                }

                return result;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Произошла ошибка: {ex.Message}. Попробуйте снова.");
            }
        }
    }

    #endregion

    #region Number Validation

    public static T ReadValidatedNumber<T>(string prompt, T min, T max, bool allowIntegers = true) where T : IComparable
    {
        while (true)
        {
            try
            {
                Console.Write(prompt);
                string input = Console.ReadLine();

                if (string.IsNullOrEmpty(input))
                {
                    Console.WriteLine("Ошибка: Ввод не может быть пустым. Попробуйте снова.");
                    continue;
                }

                T value = (T)Convert.ChangeType(input, typeof(T));

                if (value.CompareTo(min) < 0 || value.CompareTo(max) > 0)
                {
                    Console.WriteLine($"Ошибка: Значение должно быть в диапазоне от {min} до {max}. Попробуйте снова.");
                    continue;
                }

                if (!allowIntegers && value is double doubleValue)
                {
                    if (doubleValue % 1 == 0)
                    {
                        Console.WriteLine("Ошибка: Требуется дробное число, а не целое. Попробуйте снова.");
                        continue;
                    }
                }

                return value;
            }
            catch (FormatException)
            {
                Console.WriteLine($"Ошибка: Неверный формат числа. Ожидается тип {typeof(T).Name}. Попробуйте снова.");
            }
            catch (OverflowException)
            {
                Console.WriteLine("Ошибка: Число слишком большое или слишком маленькое. Попробуйте снова.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Произошла ошибка: {ex.Message}. Попробуйте снова.");
            }
        }
    }

    public static double ReadValidatedDouble(string prompt, double min, double max, bool allowIntegers = true)
    {
        return ReadValidatedNumber<double>(prompt, min, max, allowIntegers);
    }

    public static int ReadValidatedInt(string prompt, int min, int max)
    {
        return ReadValidatedNumber<int>(prompt, min, max, true);
    }

    #endregion

    #region String Validation

    public static string ReadValidatedString(string prompt, int minLength, int maxLength)
    {
        while (true)
        {
            try
            {
                Console.Write(prompt);
                string input = Console.ReadLine();

                if (input == null)
                {
                    Console.WriteLine("Ошибка: Ввод не может быть null. Попробуйте снова.");
                    continue;
                }

                if (input.Length < minLength || input.Length > maxLength)
                {
                    Console.WriteLine($"Ошибка: Длина строки должна быть от {minLength} до {maxLength} символов. Попробуйте снова.");
                    continue;
                }

                return input;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Произошла ошибка: {ex.Message}. Попробуйте снова.");
            }
        }
    }

    #endregion

    #region Enum Validation

    public static T ReadValidatedEnum<T>(string prompt) where T : struct, Enum
    {
        while (true)
        {
            try
            {
                Console.Write(prompt);
                string input = Console.ReadLine();

                if (string.IsNullOrEmpty(input))
                {
                    Console.WriteLine("Ошибка: Ввод не может быть пустым. Попробуйте снова.");
                    continue;
                }

                if (Enum.TryParse<T>(input, true, out T result) && Enum.IsDefined(typeof(T), result))
                {
                    return result;
                }
                else
                {
                    var validValues = string.Join(", ", Enum.GetValues(typeof(T)).Cast<T>());
                    Console.WriteLine($"Ошибка: Недопустимое значение. Допустимые значения: {validValues}. Попробуйте снова.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Произошла ошибка: {ex.Message}. Попробуйте снова.");
            }
        }
    }

    #endregion

    #region Boolean Validation

    public static bool ReadValidatedBoolean(string prompt, string trueValue = "да", string falseValue = "нет")
    {
        while (true)
        {
            try
            {
                Console.Write(prompt);
                string input = Console.ReadLine()?.ToLower().Trim();

                if (string.IsNullOrEmpty(input))
                {
                    Console.WriteLine("Ошибка: Ввод не может быть пустым. Попробуйте снова.");
                    continue;
                }

                if (input == trueValue.ToLower() || input == "y" || input == "yes" || input == "1" || input == "true")
                    return true;
                else if (input == falseValue.ToLower() || input == "n" || input == "no" || input == "0" || input == "false")
                    return false;
                else
                    Console.WriteLine($"Ошибка: Введите '{trueValue}' или '{falseValue}'. Попробуйте снова.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Произошла ошибка: {ex.Message}. Попробуйте снова.");
            }
        }
    }

    #endregion
}