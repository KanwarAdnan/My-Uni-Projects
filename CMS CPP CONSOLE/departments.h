#include <iostream>
#include <fstream>
#include <sstream>
#include <cstring>
#include <iomanip>

using namespace std;

class Departments
{
private:
    int count;

    struct department
    {
        int id;
        string code, description, lectures;
    };

    department deps[100];

public:
    Departments()
    {
        count = 0;
        readData();
    }

    void menu();
    void writeData();
    void readData();
    void add_department();
    void update_department();
    void delete_department();
    void search_department();
    void show_department(int id);
    void show();
    bool isIndex(int index) { return ((index >= 0) && (index < count) ? true : false); }

    bool isId(int id, int *ID)
    {
        for (int i = 0; i < count; i++)
        {
            if (id == deps[i].id)
            {
                *ID = i;
                return true;
            }
        }
        return false;
    }

    bool isCode(string code)
    {
        for (int i = 0; i < count; i++)
        {
            if (code == deps[i].code)
            {
                return false;
            }
        }
        return true;
    }

    int getCount() { return this->count; }
    void setCount(int count) { this->count = count; }
};

void Departments::menu()
{
    int temp;
    while (true)
    {
        system("cls");
        cout << "\nDEPARTMENTS TAB\n";
        cout << "\n1) ADD DEPARTMENT";
        cout << "\n2) UPDATE DEPARTMENT";
        cout << "\n3) DELETE DEPARTMENT";
        cout << "\n4) SEARCH DEPARTMENT";
        cout << "\n5) DISPLAY DEPARTMENTS";
        cout << "\n6) SAVE AND EXIT";

        cout << "\n\nSELECT AN OPTION : ";
        cin >> temp;

        cin.ignore();

        switch (temp)
        {
        case 1:
            add_department();
            writeData();
            system("cls");
            break;
        case 2:
            update_department();
            writeData();
            system("cls");
            break;
        case 3:
            delete_department();
            writeData();
            system("cls");
            break;
        case 4:
            search_department();
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

void Departments::writeData()
{
    ofstream file;
    file.open("data\\departments.csv");

    if (file.is_open())
    {

        for (int i = 0; i < count; i++)
        {
            file << deps[i].id << ",";
            file << deps[i].code << ",";
            file << deps[i].description << ",";
            file << deps[i].lectures << "\n";
        }
        file.close();
        system("cls");
    }
}

void Departments::readData()
{
    ifstream file;
    file.open("data\\departments.csv");
    string temp;

    if (file.is_open())
    {

        while (getline(file, temp))
        {
            string *ptr = split(temp);
            deps[count].id = stoi(ptr[0]);
            deps[count].code = ptr[1];
            deps[count].description = ptr[2];
            deps[count].lectures = ptr[3];
            count++;
        }
        file.close();
        system("cls");
    }
}

void Departments::search_department()
{
    system("cls");
    int id;
    int ID;
    string temp;

    cout << "\nSearch Department\n\n";

    cout << "\nEnter Department Id : ";
    cin >> id;
    cin.ignore();

    if (id <= 0)
    {
        menu();
    }

    if (isId(id, &ID))
    {
        system("cls");
        cout << "\nSearch Result\n";
        show_department(ID);
    }
    else
    {
        system("cls");
        cout << "\nCouldn't find any record !!!\n\n";
        system("pause");
        search_department();
    }
    system("pause");
}


void Departments::add_department()
{
    system("cls");

    string temp;

start:

    cout << "\nAdd Department\n\n";

Title:

    cout << "\nDepartment Title : ";
    getline(cin, temp);
    _upper(temp);

    if (temp == "")
    {
        goto Title;
    }

    // Checking If the same code already exits.

    if (isCode(temp))
    {
        deps[count].code = temp.substr(0,32);
    }
    else
    {
        system("cls");
        cout << "\nCode already present. Try another Code!";
        goto start;
    }

Desc:

    cout << "\nDepartment Short Description : ";
    getline(cin, temp);
    _upper(temp);

    if (temp == "")
    {
        goto Desc;
    }

    deps[count].description = temp.substr(0,32);

Lec:

    cout << "\nDepartment Lectures via Bar (|) Separation : ";
    getline(cin, temp);
    _upper(temp);

    if (temp == "")
    {
        goto Lec;
    }

    deps[count].lectures = temp;

    if (count == 0)
    {
        deps[0].id = 1;
    }
    else
    {
        deps[count].id = (deps[count - 1].id) + 1; // ID OF THE LAST DEP + 1
    }

    count++;

    system("cls");
    cout << "\nAdded !!!\n\n";
    system("pause");
}

void Departments::show()
{

    system("cls");
    cout << setw(8) << left << "SR.#" << setw(8) << left << "ID" << setw(40) << left << "DEPARTMENT NAME"  << setw(40) << left << "DESCRIPTION" << setw(50) << left << "LECTURES" << endl;
    cout << "---------------------------------------------------------------------------------------------------------------------------------------------" << endl;

    for (int i = 0; i < count; i++)
    {
        cout << setw(8) << left << (i+1) << setw(8) << left << deps[i].id << setw(40) << left << deps[i].code << setw(40) << left << deps[i].description << setw(50) << left << deps[i].lectures << endl;
    }

    cout << "\n";
    system("pause");
}

void Departments::show_department(int id)
{
    system("cls");
    cout << setw(8) << left << "ID" << setw(40) << left << "DEPARTMENT NAME"  << setw(40) << left << "DESCRIPTION" << setw(50) << left << "LECTURES" << endl;
    cout << "---------------------------------------------------------------------------------------------------------------------------------------------" << endl;
    cout << setw(8) << left << deps[id].id << setw(40) << left << deps[id].code << setw(40) << left << deps[id].description << setw(50) << left << deps[id].lectures << endl;
    cout << "\n\n";
}


void Departments::update_department()
{
    system("cls");
    int id, ID;
    string temp;

    cout << "\nUpdate Department\n\n";

    cout << "\nEnter Department Id : ";
    cin >> id;

    cin.ignore();

    if (id <= 0)
    {
        menu();
    }

    if (isId(id, &ID))
    {

    start:

        show_department(ID);

        cout << "\nNew Information \n";

    Title:

        cout << "\nDepartment Title : ";
        getline(cin, temp);
        _upper(temp);

        if (temp == "")
        {
            goto Title;
        }

        if (isCode(temp))
        {
            deps[ID].code = temp.substr(0,32);
        }
        else
        {
            system("cls");
            cout << "\nCode already present. Try another Code!\nUpdate Department\n\n";
            system("pause");
            goto start;
        }

    Desc:

        cout << "\nDepartment Short Description : ";
        getline(cin, temp);
        _upper(temp);

        if (temp == "")
        {
            goto Desc;
        }

        deps[ID].description = temp.substr(0,32);

    Lec:

        cout << "\nDepartment Lectures via Bar (|) Separation : ";
        getline(cin, temp);
        _upper(temp);

        if (temp == "")
        {
            goto Lec;
        }

        deps[ID].lectures = temp;

        system("cls");
        show_department(ID);
        cout << "\n";

        cout << "\nUpdated !!!\n\n";
    }

    else
    {
        system("cls");
        cout << "\nCouldn't find any record !!!\n\n";
        system("pause");
        update_department();
    }

    system("pause");
}

void Departments::delete_department()
{
    system("cls");
    char choice;

    int id, ID;
    string temp;

    cout << "\nDelete Department\n\n";

    cout << "\nEnter Department Id : ";
    cin >> id;

    cin.ignore();

    if (id <= 0)
    {
        menu();
    }

    if (isId(id, &ID))
    {
        show_department(ID);
        cout << "\nDo you really wish to delete? (y/any) : ";
        cin >> choice;
        if (choice == 'Y' || choice == 'y')
        {
            for (int i = ID; i < count; i++)
            {
                deps[i].id = deps[i + 1].id;
                deps[i].code = deps[i + 1].code;
                deps[i].description = deps[i + 1].description;
                deps[i].lectures = deps[i + 1].lectures;
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
        delete_department();
    }
    cin.ignore();
}