#include <cs50.h>
#include <math.h>
#include <stdio.h>
#include <string.h>

int every_other_digit(long cc);
int multiplyAndSum(int last_digit);
int number_of_digits(long cc);
bool isValidAmex(long cc, int numDigits);
bool isValidMaster(long cc, int numDigits);
bool isValidVisa(long cc, int numDigits);

int main(void)
{
    long cc = get_long("Insert cc number to check ");
    int sum_every_other_digit = every_other_digit(cc);
    int numDigits = number_of_digits(cc);
    bool amex = isValidAmex(cc, numDigits);
    bool master = isValidMaster(cc, numDigits);
    bool visa = isValidVisa(cc, numDigits);

    if (sum_every_other_digit % 10 != 0)
    {
        printf("INVALID\n");
        return 0;
    }
    else if (amex == true)
    {
        printf("AMEX\n");
    }
    else if (master == true)
    {
        printf("MASTERCARD\n");
    }
    else if (visa == true)
    {
        printf("VISA\n");
    }
    else
    {
        printf("INVALID\n");
        return 0;
    }
}

bool isValidVisa(long cc, int numDigits)
{
    if (numDigits == 13)
    {
        int first_digit = cc / pow(10, 12);
        if (first_digit == 4)
        {
            return true;
        }

    }
   else if (numDigits == 16)
    {
        int first_digit = cc / pow(10, 15);
        if (first_digit == 4)
        {
            return true;
        }
    }
    return false;
}



bool isValidMaster(long cc, int numDigits)
{
    int first_two_digits = cc / pow(10, 14);
    if ((numDigits == 16) && (first_two_digits > 50 && first_two_digits < 56))
    {
        return true;
    }
    else
    {
        return false;
    }
}
bool isValidAmex(long cc, int numDigits)
{
    int first_two_digits = cc / pow(10, 13);
    if ((numDigits == 15) && (first_two_digits == 34 || first_two_digits == 37))
    {
        return true;
    }
    else
    {
        return false;
    }
}

int number_of_digits(long cc)
{
    int count = 0;
    while (cc > 0)
    {
        count = count + 1;
        cc = cc / 10;
    }
    return count;
}

int every_other_digit(long cc)
{
    int sum = 0;
    bool isAlternateDigit = false;
    while (cc > 0)
    {
        if (isAlternateDigit == true)
        {
            int last_digit = cc % 10;
            int product = multiplyAndSum(last_digit);
            sum = sum + product;
        }
        else
        {
            int last_digit = cc % 10;
            sum = sum + last_digit;
        }
        isAlternateDigit = !isAlternateDigit;
        cc = cc / 10;
    }
    return sum;
}

int multiplyAndSum(int last_digit)
{
    int multiply = last_digit * 2;
    int sum = 0;
    while (multiply > 0)
    {
        int last_digit_multiply = multiply % 10;
        sum = sum + last_digit_multiply;
        multiply = multiply / 10;
    }
    return sum;
}
