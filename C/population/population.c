#include <cs50.h>
#include <stdio.h>

int main(void)
{
    const int min_pop = 9;
    int start_pop;
    int end_pop;

    do
    {
        start_pop = get_int("Please enter starting population (min = 9) ");
    }
    while (start_pop < min_pop);
    do
    {
        end_pop = get_int("Please enter end population (must be higher than the min pop set) ");
    }
    while (end_pop < start_pop);

    // TODO: Calculate number of years until we reach threshold
    int n = start_pop;
    int years = 0;
    while (n < end_pop)
    {
        n = n + n / 3 - n / 4;
        years++;
    }

    // TODO: Print number of years
    printf("Years: %i\n", years);
}
