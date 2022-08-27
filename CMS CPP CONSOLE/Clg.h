#include <iostream>
#include "csv_reader.h"
#include "departments.h"
#include "students.h"
#include "teachers.h"
#include "attendances.h"

using namespace std;

class CLG
{
private:
    string name;
    string address;
    string owner;
    string date_of_built;
    string admin_password;
    Departments dep;
    Students stu;
    Teachers tech;
    Attendance attend;

public:
    CLG()
    {
        this->name = "Govt College Of Science";
        this->address = "Wahdat Road Lahore, Pakistan";
        this->owner = "Kanwar Adnan";
        this->date_of_built = "11/11/2001";
        this->admin_password = "admin@gcs.edu.pk";
    }

    void menu();

    string getName() { return this->name; }
    string getAddress() { return this->address; }
    string getOwner() { return this->owner; }
    string getDate() { return this->date_of_built; }
    string getPassword() { return this->admin_password; }

    void setName(string name) { this->name = name; }
    void setAddress(string address) { this->address = address; }
    void setOwner(string owner) { this->owner = owner; }
    void setDate(string date) { this->date_of_built = date; }
    void setPassword(string password) { this->admin_password = password; }
};

void CLG::menu()
{
    int temp;
    while (true)
    {
        system("cls");

        cout << "\nMAIN MENU\n";
        cout << "\n1) DEPARTMENTS";
        cout << "\n2) STUDENTS";
        cout << "\n3) TEACHERS";
        cout << "\n4) ATTENDANCE";
        cout << "\n5) LOGIN";
        cout << "\n6) BACK";

        cout << "\n\nSELECT AN OPTION : ";
        cin >> temp;

        cin.ignore();

        switch (temp)
        {
        case 1:
            dep.setCount(0);
            dep.readData();
            dep.menu();
            dep.writeData();
            system("cls");
            break;
        case 2:
            stu.setCount(0);
            stu.setDepCount(0);
            stu.readDep();
            stu.readData();
            stu.menu();
            stu.writeData();
            system("cls");
            break;
        case 3:
            tech.setCount(0);
            tech.setDepCount(0);
            tech.readDep();
            tech.readData();
            tech.menu();
            tech.writeData();
            system("cls");
            break;
        case 4:
            system("cls");
            attend.setCount(0);
            attend.setDepCount(0);
            attend.setStudCount(0);
            attend.readDep();
            attend.readStud();
            attend.readData();
            attend.menu();
            attend.writeData();
            break;
        case 5:
            system("cls");
            break;
        case 6:
            stu.writeData();
            dep.writeData();
            attend.writeData();
            exit(0);
        case -1:
            exit(0);
        default:
            system("cls");
            cout << "\nKya bey salay tjy integer or alphabet me farq b ni pta ! \n\n";
            system("pause");
            exit(0);
            break;
        }
    }
}