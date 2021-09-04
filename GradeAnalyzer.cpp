#include <OpenXLSX.hpp>
#include <iostream>

using namespace std;
using namespace OpenXLSX;

constexpr unsigned int stoint(const char* str, int h = 0)
{
    return !str[h] ? 5381 : (stoint(str, h + 1) * 33) ^ str[h];
}

double gpaletter(string letter)
{
    switch (stoint(&letter[0])) {
        case stoint("A+"):
            return 4;
        case stoint("A"):
            return 4;
        case stoint("A-"):
            return 3.7;
        case stoint("B+"):
            return 3.3;
        case stoint("B"):
            return 3;
        case stoint("B-"):
            return 2.7;
        case stoint("C+"):
            return 2.3;
        case stoint("C"):
            return 2;
        case stoint("C-"):
            return 1.7;
        case stoint("D+"):
            return 1.3;
        case stoint("D"):
            return 1;
        case stoint("D-"):
            return 0.7;
        case stoint("F"):
            return 0;
        case stoint("P"):
            return -1;
        case stoint("NP"):
            return -1;
        case stoint("S"):
            return -1;
        case stoint("U"):
            return -1;
        case stoint("IP"):
            return -1;
        default:
            cout << "gpa error";
            return -99999;
            break;
    }
    return -99999;
};

int gindex(string letter)
{
    switch (stoint(&letter[0])) {
        case stoint("A+"):
            return 0;
        case stoint("A"):
            return 1;
        case stoint("A-"):
            return 2;
        case stoint("B+"):
            return 3;
        case stoint("B"):
            return 4;
        case stoint("B-"):
            return 5;
        case stoint("C+"):
            return 6;
        case stoint("C"):
            return 7;
        case stoint("C-"):
            return 8;
        case stoint("D+"):
            return 9;
        case stoint("D"):
            return 10;
        case stoint("D-"):
            return 11;
        case stoint("F"):
            return 12;
        case stoint("P"):
            return 13;
        case stoint("NP"):
            return 14;
        case stoint("S"):
            return 15;
        case stoint("U"):
            return 16;
        case stoint("IP"):
            return 17;
        default:
            cout << "geindex error " << letter << " was found ";
            return -99999;
            break;
    }
    return -99999;
};

double gpa(int ind)
{
    switch (ind) {
        case 0:
            return 4;
        case 1:
            return 4;
        case 2:
            return 3.7;
        case 3:
            return 3.3;
        case 4:
            return 3;
        case 5:
            return 2.7;
        case 6:
            return 2.3;
        case 7:
            return 2;
        case 8:
            return 1.7;
        case 9:
            return 1.3;
        case 10:
            return 1;
        case 11:
            return 0.7;
        case 12:
            return 0;
        default:
            cout << "float gpa error " << ind << "\n";
            return -99999;
            break;
    }
    return -99999;
};

string nextquarter(string current)
{
    string season = current.substr(0, current.find(' '));
    string year   = current.substr(current.find(' ')+1, current.length() - 1);
    if (season == "Fall") {
        int numyear = stoi(year) + 1;
        year        = to_string(numyear);
    }
    switch (stoint(&season[0])) {
        case stoint("Fall"):
            season = "Winter";
            break;
        case stoint("Winter"):
            season = "Spring";
            break;
        case stoint("Spring"):
            season = "Summer";
            break;
        case stoint("Summer"):
            season = "Fall";
            break;
        default:
            break;
    }
    return season + " " + year;
};

struct coursedata{
    string coursename = "";
    int grades[18] = {0};
};

int main()
{
    cout << "********************************************************************************\n";
    cout << "Grade Analyzer for UCSB grades\n";
    cout << "********************************************************************************\n";

    auto PrintCell = [](const XLCell& cell) {
        cout << "Cell type is ";

        switch (cell.valueType()) {
            case XLValueType::Empty:
                cout << "XLValueType::Empty";
                break;

            case XLValueType::Float:
                cout << "XLValueType::Float and the value is " << cell.value().get<double>() << endl;
                break;

            case XLValueType::Integer:
                cout << "XLValueType::Integer and the value is " << cell.value().get<int64_t>() << endl;
                break;

            case XLValueType::String:
                cout << "XLValueType::String and the value is " << cell.value().get<std::string>() << endl;
                break;
            default:
                cout << "Unknown";
        }
    };

    XLDocument doc;
    doc.open("C:\\Users\\14086\\Documents\\GitHub\\OpenXLSX\\examples\\grades.xlsx");
    string pre_quarter = "Summer 2015";//Enter season right before first quarter
    //i.e. To start from Fall 2015, enter Summer 2015;
    string end_quarter = "Winter 2021";
    auto wks = doc.workbook().worksheet(end_quarter);
    int row =-1;
    bool undergrad = false;
    map<string, coursedata> classroom;//course subjects, course data
    map<string, coursedata> lessons;//instructors, course data
    // Course Name-Undergraduate
    //  Instructors
    // Grades
    
    while (pre_quarter != end_quarter) {
        pre_quarter = nextquarter(pre_quarter);
        cout << pre_quarter << "\n";
        wks = doc.workbook().worksheet(pre_quarter);
        row = 2;
        while (wks.cell("A" + to_string(row)).valueType() != XLValueType::Empty){
            string classname = wks.cell("B" + to_string(row)).value().get<std::string>();
            if (classroom.find(classname) == classroom.end()) {
                coursedata cdata;
                cdata.coursename = classname;
                classroom.insert(pair<string, coursedata>(classname, cdata));
            }
            string instructor = wks.cell("C" + to_string(row)).value().get<std::string>();
            string lesson = classname + " by " + instructor;
            if (lessons.find(lesson) == lessons.end()) {
                coursedata cdata;
                cdata.coursename = lesson;
                lessons.insert(pair<string, coursedata>(lesson, cdata));
            }
            string grade = wks.cell("D" + to_string(row)).value().get<std::string>();
            int quantity = wks.cell("E" + to_string(row)).value().get<int64_t>();
            int    dummygindex = gindex(grade);
            classroom[classname].grades[gindex(grade)] += quantity;
            lessons[lesson].grades[gindex(grade)] += quantity;
            row+=1;
        }
        cout << pre_quarter << " has been processed.\n";
        
    }

    XLDocument outdoc;
    outdoc.create("C:\\Users\\14086\\Documents\\GitHub\\OpenXLSX\\examples\\UCSBGrades.xlsx");
    auto firstsheet = outdoc.workbook().worksheet("Sheet1");
    firstsheet.updateSheetName("Sheet1","Grades by Course");
    auto outwks   = firstsheet;//outdoc.workbook().worksheet("Grades by Course");
    outwks.cell("A1").value() = "UCSB Grades by Course";
    int outrow = 2;
    for (auto const& [cname, cdata] : classroom) {
        outwks.cell("A" + to_string(outrow)).value() = cname;
        float mean = 0;
        float count = 0;
        for (int a = 0; a < 13; a++) {
            mean += cdata.grades[a] * gpa(a);
            count += cdata.grades[a];
        }
        mean/=count;
        int halfcount = 0;
        int b = 0;
        while (halfcount <= count/2 && b<13) {
            halfcount += cdata.grades[b];
            b+=1;
        }
        float median = gpa(b-1);
        outwks.cell("B" + to_string(outrow)).value()   = "Mean";
        outwks.cell("C" + to_string(outrow)).value() = mean;
        outwks.cell("B" + to_string(outrow+1)).value() = "Median";
        outwks.cell("C" + to_string(outrow+1)).value() = median;
        outwks.cell("B" + to_string(outrow + 2)).value() = "A";
        outwks.cell("C" + to_string(outrow+2)).value() = cdata.grades[0];
        outwks.cell("B" + to_string(outrow + 3)).value() = "F";
        outwks.cell("C" + to_string(outrow+3)).value() = cdata.grades[11];
        outrow+=4;
    }

    outdoc.workbook().addWorksheet("Grades by Instructor");
    auto secondsheet = outdoc.workbook().worksheet("Grades by Instructor");
    auto outwks2 = secondsheet;
    outwks2.cell("A1").value() = "UCSB Grades by Instructor";
    outrow = 2;
    for (auto const& [cname, cdata] : lessons) {
        outwks2.cell("A" + to_string(outrow)).value() = cname;
        float mean = 0;
        float count = 0;
        for (int a = 0; a < 13; a++) {
            mean += cdata.grades[a] * gpa(a);
            count += cdata.grades[a];
        }
        mean /= count;
        int halfcount = 0;
        int b         = 0;
        while (halfcount <= count/2 && b<13) {
            halfcount += cdata.grades[b];
            b += 1;
        }
        float median                                     = gpa(b - 1);
        outwks2.cell("B" + to_string(outrow)).value()     = "Mean";
        outwks2.cell("C" + to_string(outrow)).value()     = mean;
        outwks2.cell("B" + to_string(outrow + 1)).value() = "Median";
        outwks2.cell("C" + to_string(outrow + 1)).value() = median;
        outwks2.cell("B" + to_string(outrow + 2)).value() = "A";
        outwks2.cell("C" + to_string(outrow + 2)).value() = cdata.grades[0];
        outwks2.cell("B" + to_string(outrow + 3)).value() = "F";
        outwks2.cell("C" + to_string(outrow + 3)).value() = cdata.grades[11];
        outrow += 4;
    }


    //for (auto const& [key, val] : symbolTable)
    /*while(wks.cell("A" + to_string(row).valueType() != XLValueType::Empty)){
        if(wks.cell("A" + to_string(row)) == "Undergraduate")
            undergrad = true;
        row++;
    }*/

    cout << "Cell A1: ";
    PrintCell(outwks.cell("A1"));

    cout << "Cell B1: ";
    PrintCell(wks.cell("B1"));

    cout << "Cell C1: ";
    PrintCell(wks.cell("C1"));

    
    doc.close();
    outdoc.save();
    outdoc.saveAs("UCSBGrades.xlsx");
    outdoc.close();

    return 0;
}
