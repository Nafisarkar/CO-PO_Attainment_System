import random
import names
import csv


def total(x, y, z, a, b, c):
    return int(x) + int(y) + int(z) + int(a) + int(b) + int(c)


def rndid():
    return (
        str(random.randint(190, 210))
        + "-"
        + str(random.randint(10, 30))
        + "-"
        + str(random.randint(14000, 14800))
    )


def midres():
    return (
        str(random.randint(0, 7))
        + ","
        + str(random.randint(0, 4))
        + ","
        + str(random.randint(0, 5))
        + ","
        + str(random.randint(0, 7))
        + ","
        + str(random.randint(0, 4))
        + ","
        + str(random.randint(0, 4))
    )


def finalres():
    return (
        str(random.randint(0, 5))
        + ","
        + str(random.randint(0, 9))
        + ","
        + str(random.randint(0, 7))
        + ","
        + str(random.randint(0, 7))
        + ","
        + str(random.randint(0, 9))
        + ","
        + str(random.randint(0, 9))
    )


x = int(input())

with open("student.csv", "w", newline="") as file:
    file = csv.writer(file)
    file.writerow(
        [
            "Name",
            "ID",
            "P1",
            "P2",
            "P3",
            "P4",
            "P5",
            "P6",
            "MID",
            "P1",
            "P2",
            "P3",
            "P4",
            "P5",
            "P6",
            "Final",
        ]
    )
    file.writerow(
        [
            "",
            "",
            "C1",
            "C2",
            "C3",
            "C4",
            "C5",
            "C6",
            "Total",
            "C1",
            "C2",
            "C3",
            "C4",
            "C5",
            "C6",
            "Total",
        ]
    )

for i in range(x):
    with open("student.csv", "a", newline="") as file:
        file = csv.writer(file)
        temp_name = names.get_full_name()
        temp_id = rndid()
        temp_a, temp_b, temp_c, temp_d, temp_e, temp_f = midres().split(",")
        temp_fa, temp_fb, temp_fc, temp_fd, temp_fe, temp_ff = finalres().split(",")
        tp = (
            temp_name,
            temp_id,
            int(temp_a),
            int(temp_b),
            int(temp_c),
            int(temp_d),
            int(temp_e),
            int(temp_f),
            int(total(temp_a, temp_b, temp_c, temp_d, temp_e, temp_f)),
            int(temp_fa),
            int(temp_fb),
            int(temp_fc),
            int(temp_fd),
            int(temp_fe),
            int(temp_ff),
            int(total(temp_fa, temp_fb, temp_fc, temp_fd, temp_fe, temp_ff)),
        )
        file.writerow(tp)
        print(names.get_full_name(), rndid(), midres(), finalres())
