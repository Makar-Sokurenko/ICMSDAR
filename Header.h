#pragma once
#include <string>
using namespace std;

class Student {
public:
    string lName;
    string fName;
    string ftName;
    string phone;  // Format: DD/MM/YYYY
    string city;
};

void copyProp(xlnt::workbook source_wb, xlnt::workbook new_wb);
void create_Sample(string subj_names[100], int subj_num, int grades_num[20]);
void load_xlsx(xlnt::workbook wb, string subj_names[100], int subj_num, int grades_num[20], Student students[100], int numStudents, double grades[100][100][100]);
void save_xlsx(xlnt::workbook wb, xlnt::workbook new_wb, string subj_names[100], int subj_num, int grades_num[20], Student students[100], int numStudents, double grades[100][100][100]);
void displayMenu();
void enterStudentData(Student students[100], int n);
void enterGradesData(Student students[100], int numStudents, double grades[100][100][100], int subj_num, int grades_num[20], string subj_names[100]);
void viewStudentData(Student students[100], int numStudents);
void viewGradesData(Student students[100], int numStudents, double grades[100][100][100], int subj_num, int grades_num[20], string subj_names[100]);

void analyzeGradesData(Student students[100], int numStudents, double grades[100][100][100], int subj_num, int grades_num[20], string subj_names[100]);
bool isValidGrade(int grade);
void clearBuffer();
