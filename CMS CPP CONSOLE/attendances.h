#include <iostream>
#include <iomanip> // for better display

using namespace std;

class Attendance
{
private:
    int count;

    struct attendance
    {
        int id, day, month, year;
        string code, subject, record, rollnos;
    };

    int dep_count;
    string deps[100] = {};

    int lectures_count;
    string lectures[100] = {};

    int stud_count;

    struct students
    {
        int id;
        string code, name;
    };

    int attendance_list[100] = {};

    students studs[100];

    attendance attends[100];

public:
    Attendance()
    {
        lectures_count = 0;
        dep_count = 0;
        stud_count = 0;
        count = 0;
        readStud();
        readDep();
        readData();
    }

    void readDep();
    void readStud();
    void show_stud();
    void menu();
    void add_attendance();
    void update_attendance();
    void delete_attendance();
    void search_attendance();

    void writeData();
    void readData();
    void show_dep();
    void show_attendance(int ID);
    void show();

    bool idWhere(int day, int month, int year, string code, string lecture, int *ID)
    {
        for (int i = 0; i < count; i++)
        {
            if ((attends[i].code == code) && (attends[i].day == day) && (attends[i].month == month) && (attends[i].year == year) && (attends[i].subject == lecture))
            {
                *ID = i;
                return true;
            }
        }
        return false;
    }

    bool isIndex(int index) { return ((index >= 0) && (index < count) ? true : false); }

    int *getStudents(string code, int *count)
    {
        int *arr = new int[100];
        int index = 0;

        for (int i = 0; i < stud_count; i++)
        {
            if (studs[i].code == code)
            {
                arr[index] = i;
                index++;
            }
        }

        *count = index;

        return arr;
    }

    bool isId(int id, int *ID)
    {
        for (int i = 0; i < count; i++)
        {
            if (id == attends[i].id)
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
            if (id == attends[i].id)
            {
                return false;
            }
        }
        return true;
    }

    int getCount() { return this->count; }
    void setCount(int count) { this->count = count; }
    void setDepCount(int dep_count) { this->dep_count = dep_count; }
    void setStudCount(int stud_count) { this->dep_count = stud_count; }
};

void Attendance::menu(){
        int temp;
    while (true)
    {
        system("cls");

        cout << "\nATTENDANCES TAB\n";
        cout << "\n1) ADD ATTENDANCE";
        cout << "\n2) UPDATE ATTENDANCE";
        cout << "\n3) DELETE ATTENDANCE";
        cout << "\n4) SEARCH ATTENDANCE";
        cout << "\n5) DISPLAY ATTENDANCES";
        cout << "\n6) SAVE AND EXIT";

        cout << "\n\nSELECT AN OPTION : ";
        cin >> temp;

        cin.ignore();

        switch (temp)
        {
        case 1:
            add_attendance();
            writeData();
            system("cls");
            break;
        case 2:
            update_attendance();
            writeData();
            system("cls");
            break;
        case 3:
            delete_attendance();
            writeData();
            system("cls");
            break;
        case 4:
            search_attendance();
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

void Attendance::show_stud()
{
    for (int i = 0; i < stud_count; i++)
    {
        cout << "\n----------------------------------------------------------";
        cout << "\nInformation of Student Roll No. : " << studs[i].id;
        cout << "\nName : " << studs[i].name;
        cout << "\nCode : " << studs[i].code;
    }
    cout << "\n\n";
    system("pause");
}

void Attendance::readStud()
{
    ifstream file;
    string temp;

    file.open("data\\students.csv");

    if (file.is_open())
    {

        while (getline(file, temp))
        {
            string *ptr = split(temp);
            studs[stud_count].id = stoi(ptr[0]);
            studs[stud_count].code = ptr[1];
            studs[stud_count].name = ptr[2];
            stud_count++;
        }
        file.close();
        system("cls");
    }
}

void Attendance::readDep()
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
void Attendance::show_dep()
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

void Attendance::add_attendance()
{
    system("cls");

    string temp;
    int ID;

start:

    cout << "\nAdd Attendance\n\n";

Day:

    cout << "\nDay (must be 2 digits) : ";
    getline(cin, temp);

    if ((temp == "") || (temp.length() < 2) || (stoi(temp) > 31))
    {
        goto Day;
    }

    attends[count].day = stoi(temp);

Month:

    cout << "\nMonth (must be 2 digits) : ";
    getline(cin, temp);

    if ((temp == "") || (temp.length() < 2) || (stoi(temp) > 12))
    {
        goto Month;
    }

    attends[count].month = stoi(temp);

Year:

    cout << "\nYear (must be 4 Digits) : ";
    getline(cin, temp);

    if ((temp == "") || (temp.length() < 4))
    {
        goto Year;
    }

    attends[count].year = stoi(temp);

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
        attends[count].code = temp.substr(0, 32);
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
        attends[count].subject = temp.substr(0, 32);
    }

    else
    {
        system("cls");
        cout << "\nSelect a valid Lecture. \n\n";
        system("pause");
        goto Lec;
    }

    // Checking Whether attendance already present or not
    int mid;

    if (idWhere(attends[count].day, attends[count].month, attends[count].year, attends[count].code, attends[count].subject, &mid))
    {
        system("cls");
        show_attendance(mid);
        cout << "\n\nAttendance Already Exists !!!\n\n";
        system("pause");
        system("cls");
        goto start;
    }

    int IDs_count = 0;

    int *IDs = getStudents(attends[count].code, &IDs_count);
    string temp2;
    string temp4 = "";
    string temp5 = "";

    system("cls");

    cout << "TAKING ATTENDANCE : CLASS : " << attends[count].code << " LECTURE : " << attends[count].subject << "\n\n";
    string temp3[IDs_count];

    cout << setw(32) << left << "STUDENT NAME" << setw(16) << left << "ID" << setw(16) << left << "STATUS" << endl;
    cout << "----------------------------------------------------------------------" << endl;

    for (int i = 0; i < IDs_count/2; i++)
    {
        cout << setw(32) << left << studs[IDs[i]].name << setw(16) << left << studs[IDs[i]].id << ": ";
        cin >> temp3[i];
        _upper(temp3[i]);
        temp4 = temp4 + "|" + temp3[i];
        temp5 = temp5 + "|" + to_string(studs[IDs[i]].id);
    }

    temp5 = temp5.substr(1);
    temp4 = temp4.substr(1);

    delete IDs;

    attends[count].record = temp4;
    attends[count].rollnos = temp5;

    if (count == 0)
    {
        attends[0].id = 1;
    }
    else
    {
        attends[count].id = (attends[count - 1].id) + 1; // ID OF THE LAST ATTENDANCE + 1
    }

    count++;

    system("cls");
    cout << "\nAdded !!!\n\n";
    system("pause");
    writeData();
}

void Attendance::show()
{
    system("cls");
    cout << setw(8) << left << "SR.#" << setw(20) << left << "DATE" << setw(40) << left << "DEPARTMENT" << setw(40) << left << "LECTURE" << setw(40) << left << "T/P/A/L" << endl;
    cout << "---------------------------------------------------------------------------------------------------------------------------------------------" << endl;

    string temp = "";

    int freq; 
    string temp2 = "";
    int t = freq;
    int a=0,p=0,l=0;

    for (int i = 0; i < count; i++)
    {
        string* ptr = split2(attends[i].record,&freq,'|');
        t = freq;
        for(int j = 0 ; j < freq ; j++){
            if (ptr[j] == "P"){
                p++;
            }
            if (ptr[j] == "A"){
                a++;
            }
            if (ptr[j] == "L"){
                l++;
            }
        }
        temp2 = to_string(freq)+ "/" + to_string(p) + "/" + to_string(a) + "/" + to_string(l);
        delete ptr;
        a=0;
        p=0;
        l=0;
        t=0;

        temp = to_string(attends[i].day) + "/" + to_string(attends[i].month) + "/" + to_string(attends[i].year);
        cout << setw(8) << left << (i + 1) << setw(20) << left << temp << setw(40) << left << attends[i].code << setw(40) << left << attends[i].subject << setw(40) << left << temp2 << endl;
    }
    cout << "\n\n";
    system("pause");
}

void Attendance::writeData()
{
    ofstream file;
    file.open("data\\attendance.csv");

    if (file.is_open())
    {

        for (int i = 0; i < count; i++)
        {
            file << attends[i].id << ",";
            file << attends[i].code << ",";
            file << attends[i].day << ",";
            file << attends[i].month << ",";
            file << attends[i].year << ",";
            file << attends[i].subject << ",";
            file << attends[i].rollnos << ",";
            file << attends[i].record << "\n";
        }
        file.close();
        system("cls");
    }
}

void Attendance::readData()
{
    ifstream file;
    string temp;

    file.open("data\\attendance.csv");

    if (file.is_open())
    {

        while (getline(file, temp))
        {
            string *ptr = split(temp);
            attends[count].id = stoi(ptr[0]);
            attends[count].code = ptr[1];
            attends[count].day = stoi(ptr[2]);
            attends[count].month = stoi(ptr[3]);
            attends[count].year = stoi(ptr[4]);
            attends[count].subject = ptr[5];
            attends[count].rollnos = ptr[6];
            attends[count].record = ptr[7];
            count++;
        }
        file.close();
        system("cls");
    }
}

void Attendance::search_attendance()
{
    system("cls");
    string temp;
    int day, month, year, ID;
    string code, lecture;

start:

    cout << "\nSearch Attendance\n\n";

Day:

    cout << "\nDay (must be 2 digits) : ";
    getline(cin, temp);

    if ((temp == "") || (temp.length() < 2) || (stoi(temp) > 31))
    {
        goto Day;
    }

    day = stoi(temp);

Month:

    cout << "\nMonth (must be 2 digits) : ";
    getline(cin, temp);

    if ((temp == "") || (temp.length() < 2) || (stoi(temp) > 12))
    {
        goto Month;
    }

    month = stoi(temp);

Year:

    cout << "\nYear (must be 4 Digits) : ";
    getline(cin, temp);

    if ((temp == "") || (temp.length() < 4))
    {
        goto Year;
    }

    year = stoi(temp);

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
        code = temp.substr(0, 32);
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
        lecture = temp.substr(0, 32);
    }

    else
    {
        system("cls");
        cout << "\nSelect a valid Lecture. \n\n";
        system("pause");
        goto Lec;
    }

    if (idWhere(day, month, year, code, lecture, &ID))
    {
        system("cls");
        show_attendance(ID);
        cout << "\n\n";
        system("pause");
    }
    else
    {
        system("cls");
        cout << "\nCouldn't find any record !!!\n\n";
        system("pause");
        search_attendance();
    }
}

void Attendance::show_attendance(int ID)
{
    cout << setw(20) << left << "DATE" << setw(40) << left << "DEPARTMENT" << setw(40) << left << "LECTURE"  << setw(40) << left << "T/P/A/L" << endl;
    cout << "---------------------------------------------------------------------------------------------------------------------------------------------" << endl;

    string temp = "";
    string temp2 = "";
    int freq=0,p=0,a=0,l=0,t=0;

    string* ptr = split2(attends[ID].record,&freq,'|');
    t = freq;
    for(int j = 0 ; j < freq ; j++){
        if (ptr[j] == "P"){
            p++;
        }
        if (ptr[j] == "A"){
            a++;
        }
        if (ptr[j] == "L"){
            l++;
        }
    }

    temp2 = to_string(freq)+ "/" + to_string(p) + "/" + to_string(a) + "/" + to_string(l);
    delete ptr;

    temp = to_string(attends[ID].day) + "/" + to_string(attends[ID].month) + "/" + to_string(attends[ID].year);
    cout << setw(20) << left << temp << setw(40) << left << attends[ID].code << setw(40) << left << attends[ID].subject << left << temp2 << endl;
}

void Attendance::delete_attendance()
{
    system("cls");
    string temp;
    int day, month, year, ID;
    string code, lecture;

start:

    cout << "\nDelete Attendance\n\n";

Day:

    cout << "\nDay (must be 2 digits) : ";
    getline(cin, temp);

    if ((temp == "") || (temp.length() < 2) || (stoi(temp) > 31))
    {
        goto Day;
    }

    day = stoi(temp);

Month:

    cout << "\nMonth (must be 2 digits) : ";
    getline(cin, temp);

    if ((temp == "") || (temp.length() < 2) || (stoi(temp) > 12))
    {
        goto Month;
    }

    month = stoi(temp);

Year:

    cout << "\nYear (must be 4 Digits) : ";
    getline(cin, temp);

    if ((temp == "") || (temp.length() < 4))
    {
        goto Year;
    }

    year = stoi(temp);

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
        code = temp.substr(0, 32);
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
        lecture = temp.substr(0, 32);
    }

    else
    {
        system("cls");
        cout << "\nSelect a valid Lecture. \n\n";
        system("pause");
        goto Lec;
    }
    char choice;
    if (idWhere(day, month, year, code, lecture, &ID))
    {
        system("cls");
        show_attendance(ID);
        cout << "\nDo you really wish to delete? (y/any) : ";
        cin >> choice;
        if (choice == 'Y' || choice == 'y')
        {
            for (int i = ID; i < count; i++)
            {
                attends[i].id = attends[i + 1].id;
                attends[i].code = attends[i + 1].code;
                attends[i].day = attends[i + 1].day;
                attends[i].month = attends[i + 1].month;
                attends[i].year = attends[i + 1].year;
                attends[i].subject = attends[i + 1].subject;
                attends[i].record = attends[i + 1].record;
                attends[i].rollnos = attends[i + 1].rollnos;
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
        delete_attendance();
    }
}

void Attendance::update_attendance()
{
    system("cls");
    string temp;
    int day, month, year, ID;
    string code, lecture;

start:

    cout << "\nUpdate Attendance\n\n";

Day:

    cout << "\nDay (must be 2 digits) : ";
    getline(cin, temp);

    if ((temp == "") || (temp.length() < 2) || (stoi(temp) > 31))
    {
        goto Day;
    }

    day = stoi(temp);

Month:

    cout << "\nMonth (must be 2 digits) : ";
    getline(cin, temp);

    if ((temp == "") || (temp.length() < 2) || (stoi(temp) > 12))
    {
        goto Month;
    }

    month = stoi(temp);

Year:

    cout << "\nYear (must be 4 Digits) : ";
    getline(cin, temp);

    if ((temp == "") || (temp.length() < 4))
    {
        goto Year;
    }

    year = stoi(temp);

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
        code = temp.substr(0, 32);
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
        lecture = temp.substr(0, 32);
    }

    else
    {
        system("cls");
        cout << "\nSelect a valid Lecture. \n\n";
        system("pause");
        goto Lec;
    }
    int roll,freq2,id;
    char status;
    if (idWhere(day, month, year, code, lecture, &ID))
    {
        system("cls");
        show_attendance(ID);
        cout << "\n\n";
        cout << "\nEnter STUDENT ID : ";
        cin >> roll;
        cin.ignore();

        string* ptr =  split2(attends[ID].record,&freq,'|');
        string* ptr2 = split2(attends[ID].rollnos,&freq2,'|');
        for(int i = 0 ; i < freq ; i++){
            if (stoi(ptr2[i]) == roll){
                id = i;
            }
        }
        cout << "\nCurrent Attendance Status : " << ptr[id] << endl << endl;
        cout << "\nNew Attendance Status (p/a/l) : ";
        cin >> status;
        string temp4;

        // Updating 
        ptr[id] = status;

        for(int i = 0 ; i < freq ; i++){
            temp4 = temp4 + "|" + ptr[i];

        }
        temp4 = temp4.substr(1);

        attends[ID].record = temp4;
        delete ptr,ptr2;
        system("cls");
        cout << "\nUpdated !!!\n\n";

        system("pause");
    }
    else
    {
        system("cls");
        cout << "\nCouldn't find any record !!!\n\n";
        system("pause");
        search_attendance();
    }
}
