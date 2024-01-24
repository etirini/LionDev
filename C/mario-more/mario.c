#include <cs50.h>
#include <stdio.h>

int main(void)
{
    int col;
    int row;
    int space;
    int height = get_int("Specify your half pyramid size ");

    for (row = 0; row < height; row++)
    {
        for (space = 0; space < height - row - 1; space++)
        {
            printf(" ");
        }
        for (col = 0; col <= row; col++)
        {
            printf("#");
        }

        printf("  ");

        for (col = 0; col <= row; col++)
        {
            printf("#");
        }
        printf("\n");
    }

}
