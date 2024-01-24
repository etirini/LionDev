// Write a function to replace vowels with numbers
// Get practice with strings
// Get practice with command line
// Get practice with switch

#include <cs50.h>
#include <ctype.h>
#include <stdio.h>
#include <string.h>

char replace(char c);

int main(int argc, string argv[])
{
    int arglen = 0;
    string s = argv[1];

    if (argc != 2)
    {
        printf("Please enter only one argument\n");
        return 1;
    }
    else
    {
        arglen = strlen(argv[1]);
        for (int i = 0; i < arglen; i++)
        {
            printf("%c", replace(s[i]));
        }
        printf("\n");
    }
}

char replace(char c)
{
    if (c == 'A' || c == 'E' || c == 'I' || c == 'O')
    {
        c = tolower(c);
    }
    switch (c)
    {
        case 'a':
            c = '6';
            return c;
            break;
        case 'e':
            c = '3';
            return c;
            break;
        case 'i':
            c = '1';
            return c;
            break;
        case 'o':
            c = '0';
            return c;
            break;
        default:
            return c;
    }

    return 1;
}
