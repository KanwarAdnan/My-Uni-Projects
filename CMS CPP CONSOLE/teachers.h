#include <iostream>
#include <fstream>
#include <sstream>
#include <cstring>
#include <iomanip>

using namespace std;

class Teachers
{
private:
    int count;

    struct teachers
    {
        int id;
        string name, code, subject, email, password;
    };

    int dep_count;
    string deps[100] = {};
    int lectures_count;
    string lectures[100] = {};

    teachers teachs[100];

public:
    Teachers()
    {
        lectures_count = 0;
        dep_count = 0;
        count = 0;
        readData();
        readDep();
    }

    void readDep();
    void menu();
    void add_teacher();
    void update_teacher();
    void delete_teacher();
    void search_teacher();

    void writeData();
    void readData();
    void show_dep();
    void show_teacher(int id);
    void show();
    bool isIndex(int index) { return ((index >= 0) && (index < count) ? true : false); }

    bool isId(int id, int *ID)
    {
        for (int i = 0; i < count; i++)
        {
            if (id == teachs[i].id)
            {
                *ID = i;
                return true;
            }
        }
        return false;
    }

    bool isValidDep(string dep, int *ID)
    {
        for (int i = 0; i < dep_count; i++)
        {
            if (deps[i] == dep)
            {
                *ID = i;
                return true;
            }
        }
        return false;
    }

    bool isValidLec(string *arr, string lec, int freq)
    {
        for (int i = 0; i < freq; i++)
        {
            if (arr[i] == lec)
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
            if (id == teachs[i].id)
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

void Teachers::writeData()
{
    ofstream file;
    file.open("data\\teachers.csv");

    if (file.is_open())
    {

        for (int i = 0; i < count; i++)
        {
            file << teachs[i].id << ",";
            file << teachs[i].code << ",";
            file << teachs[i].name << ",";
            file << teachs[i].subject << ",";
            file << teachs[i].email << ",";
            file << teachs[i].password << "\n";
        }
        file.close();
        system("cls");
    }
}

void Teachers::readData()
{
    ifstream file;
    string temp;

    file.open("data\\teachers.csv");

    if (file.is_open())
    {

        while (getline(file, temp))
        {
            string *ptr = split(temp);
            teachs[count].id = stoi(ptr[0]);
            teachs[count].code = ptr[1];
            teachs[count].name = ptr[2];
            teachs[count].subject = ptr[3];
            teachs[count].email = ptr[4];
            teachs[count].password = ptr[5];
            count++;
        }
        file.close();
        system("cls");
    }
}

void Teachers::readDep()
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
            lectures[lectures_count] = ptr[3];
            lectures_count++;
            dep_count++;
        }
        file.close();
    }
}

void Teachers::show_dep()
{
    for (int i = 0; i < dep_count; i++)
    {
        cout << deps[i] << endl;
    }

    for (int i = 0; i < lectures_count; i++)
    {
        cout << lectures[i] << endl;
    }
}

void Teachers::add_teacher()
{
    system("cls");

    string temp;
    int ID;

start:

    cout << "\nAdd Teacher\n\n";

Name:

    cout << "\nTeacher Name : ";
    getline(cin, temp);
    _upper(temp);

    if (temp == "")
    {
        goto Name;
    }

    teachs[count].name = temp.substr(0,32);

Email:

    cout << "\nTeacher Email : ";
    getline(cin, temp);
    _upper(temp);

    if (temp == "")
    {
        goto Email;
    }

    teachs[count].email = temp.substr(0,32);

Pass:

    cout << "\nAccount Password : ";
    getline(cin, temp);

    if (temp == "")
    {
        goto Pass;
    }

    teachs[count].password = temp;

Dep:

    system("cls");

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
        goto Dep;
    }

    if (isValidDep(temp, &ID))
    {
        teachs[count].code = temp.substr(0,32);
    }

    else
    {
        system("cls");
        cout << "\nSelect a valid Department Name. \n\n";
        system("pause");
        goto Dep;
    }

Lec:

    system("cls");
    int freq = 0;

    cout << "\nSelect Lecture : \n";
    string one_lecture = lectures[ID];
    string *ptr = split2(one_lecture, &freq, '|');

    for (int i = 0; i < freq; i++)
    {
        cout << "\n\t"
             << (i + 1) << " : " << ptr[i];
    }

    cout << "\n\nEnter Exact Lecture Name : ";
    getline(cin, temp);
    _upper(temp);

    if (temp == "")
    {
        goto Lec;
    }

    if (isValidLec(ptr, temp, freq))
    {
        teachs[count].subject = temp.substr(0,32);
    }

    else
    {
        system("cls");
        cout << "\nSelect a valid Lecture. \n\n";
        system("pause");
        goto Lec;
    }

    if (count == 0)
    {
        teachs[0].id = 1;
    }
    else
    {
        teachs[count].id = (teachs[count - 1].id) + 1; // ID OF THE LAST TEACHER + 1
    }

    count++;

    system("cls");
    cout << "\nAdded !!!\n\n";
    system("pause");
}

void Teachers::show()
{

    system("cls");
    cout << setw(8) << left << "SR.#" << setw(8) << left << "ID" << setw(32) << left << "TEACHER NAME" << setw(32) << left << "DEPARTMENT" << setw(32) << left << "LECTURE" << setw(32) << left << "EMAIL" << endl;
    cout << "---------------------------------------------------------------------------------------------------------------------------------------------" << endl;

    for (int i = 0; i < count; i++)
    {
        cout << setw(8) << left << (i+1) << setw(8) << left << teachs[i].id << setw(32) << left << teachs[i].name << setw(32) << left << teachs[i].code << setw(32) << left << teachs[i].subject << setw(32) << left << teachs[i].email << endl;
    }

    cout << "\n\n";
    system("pause");
}

void Teachers::show_teacher(int id)
{
    system("cls");
    cout << setw(8) << left << "ID" << setw(32) << left << "TEACHER NAME" << setw(32) << left << "DEPARTMENT" << setw(32) << left << "LECTURE" << setw(32) << left << "EMAIL" << endl;
    cout << "---------------------------------------------------------------------------------------------------------------------------------------------" << endl;
    cout << setw(8) << left << teachs[id].id << setw(32) << left << teachs[id].name << setw(32) << left << teachs[id].code << setw(32) << left << teachs[id].subject << setw(32) << left << teachs[id].email << endl;
    cout << "\n\n";
}

void Teachers::delete_teacher()
{
    system("cls");
    char choice;

    int id, ID;
    string temp;

    cout << "\nDelete Teacher\n\n";

    cout << "\nEnter Teacher Id : ";
    cin >> id;

    if (id <= 0)
    {
        menu();
    }

    cin.ignore();

    if (isId(id, &ID))
    {
        show_teacher(ID);
        cout << "\nDo you really wish to delete? (y/any) : ";
        cin >> choice;
        if (choice == 'Y' || choice == 'y')
        {
            for (int i = ID; i < count; i++)
            {
                teachs[i].id = teachs[i + 1].id;
                teachs[i].code = teachs[i + 1].code;
                teachs[i].name = teachs[i + 1].name;
                teachs[i].email = teachs[i + 1].email;
                teachs[i].password = teachs[i + 1].password;
                teachs[i].subject = teachs[i + 1].subject;
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
        delete_teacher();
    }
    cin.ignore();
}

void Teachers::search_teacher()
{
    system("cls");
    int id;
    int ID;
    string temp;

    cout << "\nSearch Teacher\n\n";

    cout << "\nEnter Teacher Id : ";
    cin >> id;
    cin.ignore();

    if (id <= 0)
    {
        menu();
    }

    if (isId(id, &ID))
    {
        system("cls");
        show_teacher(ID);
    }
    else
    {
        system("cls");
        cout << "\nCouldn't find any record !!!\n\n";
        system("pause");
        search_teacher();
    }
    cout << "\n\n";
    system("pause");
}

void Teachers::update_teacher()
{
    system("cls");
    int id, ID;
    string temp;

    cout << "\nUpdate Teacher\n\n";

    cout << "\nEnter Teacher Id : ";
    cin >> id;

    cin.ignore();

    if (id <= 0)
    {
        menu();
    }

    if (isId(id, &ID))
    {

    start:

        show_teacher(ID);

        cout << "\nNew Information \n";

    Name:

        cout << "\nTeacher Name : ";
        getline(cin, temp);
        _upper(temp);

        if (temp == "")
        {
            goto Name;
        }

        teachs[ID].name = temp.substr(0,32);

    Email:

        cout << "\nTeacher Email : ";
        getline(cin, temp);
        _upper(temp);

        if (temp == "")
        {
            goto Email;
        }

        teachs[ID].email = temp.substr(0,32);

    Pass:

        cout << "\nAccount Password : ";
        getline(cin, temp);

        if (temp == "")
        {
            goto Pass;
        }

        teachs[ID].password = temp;

    Dep:

        system("cls");

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
            goto Dep;
        }

        int UID = 0;

        if (isValidDep(temp, &UID))
        {
            teachs[ID].code = temp.substr(0,32);
        }

        else
        {
            system("cls");
            cout << "\nSelect a valid Department Name. \n\n";
            system("pause");
            goto Dep;
        }

    Lec:

        system("cls");
        int freq = 0;

        cout << "\nSelect Lecture : \n";
        string one_lecture = lectures[UID];
        string *ptr = split2(one_lecture, &freq, '|');

        for (int i = 0; i < freq; i++)
        {
            cout << "\n\t"
                 << (i + 1) << " : " << ptr[i];
        }

        cout << "\n\nEnter Exact Lecture Name : ";
        getline(cin, temp);
        _upper(temp);

        if (temp == "")
        {
            goto Lec;
        }

        if (isValidLec(ptr, temp, freq))
        {
            teachs[ID].subject = temp.substr(0,32);
        }

        else
        {
            system("cls");
            cout << "\nSelect a valid Lecture. \n\n";
            system("pause");
            goto Lec;
        }
    }
    system("cls");
    show_teacher(ID);
    cout << "\n";

    cout << "\nUpdated !!!\n\n";
}

void Teachers::menu()
{
    int temp;
    while (true)
    {
        system("cls");

        cout << "\nTEACHERS TAB\n";
        cout << "\n1) ADD TEACHER";
        cout << "\n2) UPDATE TEACHER";
        cout << "\n3) DELETE TEACHER";
        cout << "\n4) SEARCH TEACHER";
        cout << "\n5) DISPLAY TEACHERS";
        cout << "\n6) SAVE AND EXIT";

        cout << "\n\nSELECT AN OPTION : ";
        cin >> temp;

        cin.ignore();

        switch (temp)
        {
        case 1:
            add_teacher();
            writeData();
            system("cls");
            break;
        case 2:
            update_teacher();
            writeData();
            system("cls");
            break;
        case 3:
            delete_teacher();
            writeData();
            system("cls");
            break;
        case 4:
            search_teacher();
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
