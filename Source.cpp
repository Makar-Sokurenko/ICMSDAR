#define _CRT_SECURE_NO_WARNINGS
#include <xlnt/xlnt.hpp>
#include <stdio.h>
#include <iostream>
#include <vector>
#include <string>
#include <limits>
#include <cctype>
#include <sstream>
#include <locale.h>
#include <Windows.h>
#include "Header.h"
using namespace std;

string subj_names[100];
int subj_num, grades_num[20];
Student students[100];
double grades[100][100][100] = { 0 };
int numStudents;

int main() {
    // ���� ������ ������� ������� � ���������� ��������� ����
    //setlocale(LC_ALL, "Ukr"); 
    setlocale(LC_ALL, "uk_UA.UTF-8");  // Set Ukrainian locale
    SetConsoleOutputCP(CP_UTF8);
    SetConsoleCP(CP_UTF8);
    numStudents = 12;
    subj_num = 5;
    subj_names[0] = u8"����������� ����'�����";
    grades_num[0] = 11;
    subj_names[1] = u8"����'������ �����������";
    grades_num[1] = 10;
    subj_names[2] = u8"����� �� ��������� �����";
    grades_num[2] = 15;
    subj_names[3] = u8"����� ���������� � ���������";
    grades_num[3] = 11;
    subj_names[4] = u8"�������㳿 �������������";
    grades_num[4] = 17;

    string file_name = u8"table.xlsx";
    // ³������ ������� ����� Excel
    xlnt::workbook wb;
    xlnt::workbook new_wb;
    wb.load(file_name);
    load_xlsx(wb, subj_names, subj_num, grades_num, students, numStudents, grades);

    while (true)
    {
        int c = 0;
        while (c < 1 || c > 8)
        {
            displayMenu();
            cin >> c;
            clearBuffer();
        }

        int n;
        switch (c)
        {
        case 1:
            cout << u8"������ ����� �������� (max - " << numStudents << ") ";
            cin >> n;
            n--;
            clearBuffer();
            enterStudentData(students, n);
            break;
        case 2:
            cout << u8"������ ����� �������� (max - " << numStudents << ") ";
            cin >> n;
            n--;
            clearBuffer();
            enterGradesData(students, n, grades, subj_num, grades_num, subj_names);
            break;
        case 3:
            viewStudentData(students, numStudents);
            break;
        case 4:
            viewGradesData(students, numStudents, grades, subj_num, grades_num, subj_names);
            break;
        case 5:
            analyzeGradesData(students, numStudents, grades, subj_num, grades_num, subj_names);
            break;
        case 6:
            save_xlsx(wb, new_wb, subj_names, subj_num, grades_num, students, numStudents, grades);
            break;
        case 7:
            create_Sample(subj_names, subj_num, grades_num);
            break;
        case 8:
            return 0;
        }
    }
}