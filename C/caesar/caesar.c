#include <cs50.h>
#include <stdio.h>
#include <string.h>
#include <ctype.h>
#include <stdlib.h>

bool k_is_digit(string usr_k);
string enc_text(int k);

int main(int argc, string argv[])
{
    if (argc != 2 || !k_is_digit(argv[1]))
    {
        printf("Usage ./caesar key\n");
        return 1;
    }
    else
    {
        int k = atoi(argv[1]);
        string ciphertext = enc_text(k);
        printf("ciphertext: %s\n", ciphertext);
        return 0;
    }
}

bool k_is_digit(string usr_k)
{
    int k_len = strlen(usr_k);
    for (int i = 0; i < k_len; i++)
    {
        if (!isdigit(usr_k[i]))
        {
            return false;
        }
    }
    return true;
}

string enc_text(int k)
{
   string plaintext = get_string("plaintext: ");
   string c = plaintext;
   for (int j = 0; j < strlen(plaintext); j++)
   {
        if (isupper(plaintext[j]))
        {
            c[j] = ((plaintext[j] - 65 + k) % 26 + 65);
        }
        else if (islower(plaintext[j]))
        {
             c[j] = ((plaintext[j] - 97 + k) % 26 + 97);
        }
        else
        {
             c[j] = (plaintext[j]);
        }
   }
    return c;
}
