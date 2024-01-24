#include <cs50.h>
#include <stdio.h>
#include <ctype.h>
#include <string.h>
#include <math.h>

int count_letters(string usr_text, int usr_text_len);
int count_words(string usr_text, int usr_text_len);
int count_sents(string usr_text, int usr_text_len);
int get_grade(int letters, int words, int sents);


int main(void)
{
    string usr_text = get_string("Text: ");
    int usr_text_len = strlen(usr_text);



    int letters = count_letters(usr_text, usr_text_len);
    int words = count_words(usr_text, usr_text_len);
    int sents = count_sents(usr_text, usr_text_len);
    int g_index = get_grade(letters, words, sents);
    /*
    printf("%i letters\n", count_letters(usr_text, usr_text_len));
    printf("%i words\n", count_words(usr_text, usr_text_len));
    printf("%i Sentences\n", count_sents(usr_text, usr_text_len));
    */

    if (g_index < 1)
    {
        printf("Before Grade 1\n");
    }
    else if (g_index > 16)
    {
        printf("Grade 16+\n");
    }
    else
    {
        printf("Grade %i\n", g_index);
    }
}

int count_letters(string usr_text, int usr_text_len)
{
    int only_letters = 0;
    for (int i = 0; i < usr_text_len; i++)
    {
        if (isalpha(usr_text[i]))
        {
            only_letters++;
        }
    }
    return only_letters;
}

int count_words(string usr_text, int usr_text_len)
{
    int count_spaces = 1;
    for (int i = 0; i < usr_text_len; i++)
    {
        if (isblank(usr_text[i]))
        {
            count_spaces++;
        }
    }
    return count_spaces;
}

int count_sents(string usr_text, int usr_text_len)
{
    int count_stops = 0;
    for (int i = 0; i < usr_text_len; i++)
    {
        if ((usr_text[i] == 33) || (usr_text[i] == 46) || (usr_text[i] == 63))
        {
            count_stops++;
        }
    }
    return count_stops;
}

int get_grade(int letters, int words, int sents)
{
    float L = (float) letters / (float) words * 100;
    float S = (float) sents / (float) words * 100;
    int index = round(0.0588 * L - 0.296 * S - 15.8);
    return index;
}

