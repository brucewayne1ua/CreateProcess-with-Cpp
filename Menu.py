import subprocess
import sys

while True:
    print("Select function")
    print("1 - Open Excel")
    print("2 - Open Word")
    print("3 - Open Access")
    print("4 - Close Menu")
    Select = input("Enter your choice: ")  # Получаем выбор пользователя

    if Select == "1":
        subprocess.run(["./OpenExcel"])
    elif Select == "2":
        subprocess.run(["./OpenWord"])
    elif Select == "3":
        subprocess.run(["./OpenAccess"])
    elif Select == "4":
        sys.exit()  # Завершаем выполнение программы
    else:
        print("Invalid choice. Please select a valid option.")