#include <iostream>
#include <fstream>
#include <sstream>
#include <cstring>
#include <iomanip>

using namespace std;

class Students
{
private:
    int count;

    struct students
    {
        int id;
        string name, code;
    };

    int dep_count;
    string deps[100] = {};

    students studs[100];

public:
    Students()
    {
        dep_count = 0;
        count = 0;
        readDep();
        readData();
    }

    void menu();
    void add_student();
    void update_student();
    void delete_student();
    void search_student();

    void writeData();
    void readDep();
    void readData();
    void show_dep();
    void show_student(int id);
    void show();
    bool isIndex(int index) { return ((index >= 0) && (index < count) ? true : false); }

    bool isId(int id, int *ID)
    {
        for (int i = 0; i < count; i++)
        {
            if (id == studs[i].id)
            {
                *ID = i;
                return true;
            }
        }
        return false;
    }

    bool isValidDep(string dep)
    {
        for (int i = 0; i < dep_count; i++)
        {
            if (deps[i] == dep)
            {
                return true;
            }
        }
        return false;
    }

    bool isID(int id)
    {
        for (int i = 0; i < count; i++)
        {
            if (id == studs[i].id)
            {
                return false;
            }
        }
        return true;
    }

    int getCount() { return this->count; }
    void setCount(int count) { this->count = count; }
    void setDepCount(int dep_count) { this->dep_count = dep_count; }
};

void Students::writeData()
{
    ofstream file;
    file.open("data\\students.csv");

    if (file.is_open())
    {

        for (int i = 0; i < count; i++)
        {
            file << studs[i].id << ",";
            file << studs[i].code << ",";
            file << studs[i].name << "\n";
        }
        file.close();
        system("cls");
    }
}

void Students::readData()
{
    ifstream file;
    string temp;

    file.open("data\\students.csv");

    if (file.is_open())
    {

        while (getline(file, temp))
        {
            string *ptr = split(temp);
            studs[count].id = stoi(ptr[0]);
            studs[count].code = ptr[1];
            studs[count].name = ptr[2];
            count++;
        }
        file.close();
        system("cls");
    }
}

void Students::add_student()
{
    string temp;

Name:
    system("cls");

    cout << "\nAdd Student\n\n";

    cout << "\nSelect Department : \n";
    for (int i = 0; i < dep_count; i++)
    {
        cout << "\n\t"
             << (i + 1) << " : " << deps[i];
    }

    cout << "\n\nEnter Exact Department Name : ";
    getline(cin, temp);
    _upper(temp);

    if (temp == "")
    {
        goto Name;
    }

    if (isValidDep(temp))
    {
        studs[count].code = temp.substr(0,32);
    }

    else
    {
        system("cls");
        cout << "\nSelect a valid Department Name. \n\n";
        system("pause");
        goto Name;
    }

    temp = "";

StudentName:

    cout << "\nEnter Student Name : ";
    getline(cin, temp);
    _upper(temp);

    if (temp == "")
    {
        goto StudentName;
    }

    studs[count].name = temp.substr(0,32);

    studs[count].id = (studs[count - 1].id) + 1; // ID OF THE LAST DEP + 1
    count++;

    system("cls");
    cout << "\nAdded !!!\n\n";
    system("pause");

}

void Students::search_student()
{
    system("cls");
    int id;
    int ID;
    string temp;

    cout << "\nSearch Student\n\n";

    cout << "\nEnter Student Id : ";
    cin >> id;
    cin.ignore();

    if (id <= 0)
    {
        menu();
    }

    if (isId(id, &ID))
    {
        system("cls");
        show_student(ID);
    }
    else
    {
        system("cls");
        cout << "\nCouldn't find any record !!!\n\n";
        system("pause");
        search_student();
    }
    cout << "\n\n";
    system("pause");
}

void Students::menu()
{
    int temp;
    while (true)
    {
        system("cls");

        cout << "\nSTUDENTS TAB\n";
        cout << "\n1) ADD STUDENT";
        cout << "\n2) UPDATE STUDENT";
        cout << "\n3) DELETE STUDENT";
        cout << "\n4) SEARCH STUDENT";
        cout << "\n5) DISPLAY STUDENTS";
        cout << "\n6) SAVE AND EXIT";

        cout << "\n\nSELECT AN OPTION : ";
        cin >> temp;

        cin.ignore();

        switch (temp)
        {
        case 1:
            add_student();
            writeData();
            system("cls");
            break;
        case 2:
            update_student();
            writeData();
            system("cls");
            break;
        case 3:
            delete_student();
            writeData();
            system("cls");
            break;
        case 4:
            search_student();
            system("cls");
            break;
        case 5:
            show();
            system("cls");
            break;
        case 6:
            system("cls");
            writeData();
            goto End;
        case -1:
            goto End;
        default:
            system("cls");
            cout << "\nKya bey salay tjy integer or alphabet me farq b ni pta ! \n\n";
            system("pause");
            exit(0);
        }
    }
End:
    cout << "\n";
}

void Students::update_student()
{
    system("cls");
    int id, ID;
    string temp;

    cout << "\nUpdate Student\n\n";

    cout << "\nEnter Student Id : ";
    cin >> id;

    cin.ignore();

    if (id <= 0)
    {
        menu();
    }

    if (isId(id, &ID))
    {

    start:

        show_student(ID);

        cout << "\nNew Information \n";

    Name:

        cout << "\nStudent Name : ";
        getline(cin, temp);
        _upper(temp);

        if (temp == "")
        {
            goto Name;
        }

        studs[ID].name = temp.substr(0,32);

    Code:

        cout << "\nSelect Department : \n";
        for (int i = 0; i < dep_count; i++)
        {
            cout << "\n\t"
                 << i << " : " << deps[i];
        }

        cout << "\n\nEnter Exact Department Name : ";
        getline(cin, temp);
        _upper(temp);

        if (temp == "")
        {
            goto Code;
        }

        if (isValidDep(temp))
        {
            studs[ID].code = temp.substr(0,32);
        }

        else
        {
            system("cls");
            goto Code;
        }

        system("cls");
        show_student(ID);
        cout << "\n";

        cout << "\nUpdated !!!\n\n";
    }

    else
    {
        system("cls");
        cout << "\nCouldn't find any record !!!\n\n";
        system("pause");
        update_student();
    }

    system("pause");
}

void Students::show()
{
    system("cls");
    cout << setw(8) << left << "SR.#" << setw(8) << left << "ID" << setw(40) << left << "STUDENT NAME" << setw(40) << left << "DEPARTMENT" << endl;
    cout << "--------------------------------------------------------------------------------" << endl;
    for (int i = 0; i < count; i++)
    {
        cout << setw(8) << left << (i+1) << setw(8) << left << studs[i].id << setw(40) << left << studs[i].name << setw(40) << left << studs[i].code << endl;
    }
    cout << "\n\n";
    system("pause");
}

void Students::show_student(int id)
{
    cout << setw(8) << left << "ID" << setw(40) << left << "STUDENT NAME" << setw(40) << left << "DEPARTMENT" << endl;
    cout << "--------------------------------------------------------------------------------" << endl;
    cout << setw(8) << left << studs[id].id << setw(40) << left << studs[id].name << setw(40) << left << studs[id].code << endl;
}

void Students::delete_student()
{
    system("cls");
    char choice;

    int id, ID;
    string temp;

    cout << "\nDelete Student\n\n";

    cout << "\nEnter Student Id : ";
    cin >> id;

    if (id <= 0)
    {
        menu();
    }

    cin.ignore();

    if (isId(id, &ID))
    {
        show_student(ID);
        cout << "\nDo you really wish to delete? (y/any) : ";
        cin >> choice;
        if (choice == 'Y' || choice == 'y')
        {
            for (int i = ID; i < count; i++)
            {
                studs[i].id = studs[i + 1].id;
                studs[i].code = studs[i + 1].code;
                studs[i].name = studs[i + 1].name;
            }
            system("cls");
            cout << "\nDeleted !!!\n\n";
            count--;
            system("pause");
        }
    }
    else
    {
        system("cls");
        cout << "\nCouldn't find any record !!!\n\n";
        system("pause");
        delete_student();
    }
    cin.ignore();
}

void Students::readDep()
{
    ifstream file;
    file.open("data\\departments.csv");
    string temp;

    if (file.is_open())
    {

        while (getline(file, temp))
        {
            string *ptr = split(temp);
            deps[dep_count] = ptr[1];
            dep_count++;
        }
        file.close();
    }
}

void Students::show_dep()
{
    for (int i = 0; i < dep_count; i++)
    {
        cout << deps[i] << endl;
    }
}