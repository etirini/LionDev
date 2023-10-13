for row in range(1,12):
    for _ in range(-row):
        print("#", end="")

    for _ in range(row):
        print("#", end="")
    print("")
