// Check that a password has at least one lowercase letter, uppercase letter, number and symbol
// Practice iterating through a string
// Practice using the ctype library
#include <cs50.h>
#include <ctype.h>
#include <stdio.h>
#include <string.h>

bool valid(string password, int passlen);
bool length(string password, int passlen);
bool symbols(string password, int passlen);
bool upanddown(string password, int passlen);
bool numbers(string password, int passlen);

int main(void)
{
    string password = get_string("Enter your password: ");

    int passlen = strlen(password);

    if (valid(password, passlen))
    {
        printf("Your password is valid!\n");
    }
    else
    {
        printf("Your password needs at least one uppercase letter, lowercase letter, number and symbol\n");
    }
}

bool valid(string password, int passlen)
{
    if (length(password, passlen) == true && symbols(password, passlen) == true && upanddown(password, passlen) == true &&
        numbers(password, passlen) == true)
    {
        return true;
    }
    else
    {
        return false;
    }
}

bool length(string password, int passlen)
{

    const int VAL_PASS = 8;

    if (passlen >= VAL_PASS)
    {
        return true;
    }
    return false;
}

bool symbols(string password, int passlen)
{
    for (int i = 0; i < passlen; i++)
    {
        if ((password[i] >= 33 && password[i] <= 47) || (password[i] >= 58 && password[i] <= 64) ||
            (password[i] >= 91 && password[i] <= 96) || (password[i] >= 123 && password[i] <= 126))
        {
            return true;
        }
    }
    return false;
}

bool upanddown(string password, int passlen)
{
    int up_val = 0;
    int down_val = 0;
    for (int i = 0; i < passlen; i++)
    {
        if (isupper(password[i]))
        {
            up_val++;
        }
        else if (islower(password[i]))
        {
            down_val++;
        }
    }
    if (up_val > 0 && down_val > 0)
    {
        return true;
    }
    return false;
}
bool numbers(string password, int passlen)
{
    for (int i = 0; i < passlen; i++)
    {
        if ((password[i] >= 48 && password[i] <= 57))
        {
            return true;
        }
    }
    return false;
}
