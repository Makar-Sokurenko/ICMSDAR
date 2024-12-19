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
#include <utf8.h>
#include "Header.h"
using namespace std;

void displayMenu() {
    cout << u8"\n����:\n";
    cout << u8"1. ������ ���������� ��� �������� (ҳ���� ��������� ����)\n";
    cout << u8"2. ������ ���������� ��� ����\n";
    cout << u8"3. ������� ���������� ��� ��������\n";
    cout << u8"4. ������� ���������� ��� ����\n";
    cout << u8"5. ����� ���������� ��� ����\n";
    cout << u8"6. �������� ���� �� xlsx �����\n";
    cout << u8"7. �������� ������ xlsx �����\n";
    cout << u8"8. Exit\n";
}
#include <locale>
#include <codecvt>
void enterStudentData(Student students[100], int n) {
    // Full Name
    cout << u8"�����: ";
    getline(cin, students[n].lName);
    // Full Name
    cout << u8"��'�: ";
    getline(cin, students[n].fName);
    // Full Name
    cout << u8"�� �������: ";
    getline(cin, students[n].ftName);
    // Full Name
    cout << u8"�������: ";
    getline(cin, students[n].phone);
    // Full Name
    cout << u8"̳���: ";
    getline(cin, students[n].city);
}

void enterGradesData(Student students[100], int n, double grades[100][100][100], int subj_num, int grades_num[20], string subj_names[100]) {

    cout << u8"\n������ ���� �������� " << n << " (" << students[n].lName << ' ' << students[n].fName << ' ' << students[n].ftName << "): \n";
    for (int j = 0; j < subj_num; j++) {
        cout << u8"������� - " << subj_names[j] << u8"\nʳ������ ���� - " << grades_num[j] - 1 << '\n';
        grades[n][j][grades_num[j] - 1] = 0;
        for (int i2 = 0; i2 < grades_num[j] - 1; i2++)
        {
            cin >> grades[n][j][i2];
            grades[n][j][grades_num[j] - 1] += grades[n][j][i2];
        }
    }
}

int utf8_length(const string& str) {
    return utf8::distance(str.begin(), str.end());
}

void viewStudentData(Student students[100], int numStudents) {
    cout << "---------------------------------------------------------------------------------------------------------------------------------------------------------\n";
    cout << u8"| � | �����                       | ��'�                          | �� �������                   |����� ��������| ̳���                              |\n";
    cout << "---------------------------------------------------------------------------------------------------------------------------------------------------------\n";
    //cout << "|Student No| Last Name| First Name | Father Name | Phone | City |\n";

    for (int i = 0; i < numStudents; i++) {
        /*cout << "| " << i + 1;
        cout << "| " << students[i].lName;
        cout << "| " << students[i].fName;
        cout << "| " << students[i].ftName;
        cout << "| " << students[i].phone;
        cout << "| " << students[i].city;
        cout << "|\n";
        */
        cout << "| " << i + 1;
        cout << string(2 - to_string(i + 1).length(), ' ') << "| " << students[i].lName;
        cout << string(30 - utf8_length(students[i].lName), ' ') << "| " << students[i].fName;
        cout << string(30 - utf8_length(students[i].fName), ' ') << "| " << students[i].ftName;
        cout << string(30 - utf8_length(students[i].ftName), ' ') << "| " << students[i].phone;
        cout << string(13 - utf8_length(students[i].phone), ' ') << "| " << students[i].city;
        cout << string(35 - utf8_length(students[i].city), ' ') << "|\n";
        
    }
    cout << "---------------------------------------------------------------------------------------------------------------------------------------------------------\n";
}


void viewGradesData(Student students[100], int numStudents, double grades[100][100][100], int subj_num, int grades_num[20], string subj_names[100]) {
    cout << "\nGrades Data:\n";

    // Print table rows for each student
    for (int i = 0; i < numStudents; i++) {
        cout << u8"-----------------------------------------------------------------------------------------------------\n";
        cout << u8"| � | �����                       | ��'�                          | �� �������                   |\n";
        cout << u8"-----------------------------------------------------------------------------------------------------\n";
        cout << u8"| " << i + 1;
        cout << string(2 - to_string(i + 1).length(), ' ') << "| " << students[i].lName;
        cout << string(30 - utf8_length(students[i].lName), ' ') << "| " << students[i].fName;
        cout << string(30 - utf8_length(students[i].fName), ' ') << "| " << students[i].ftName;
        cout << string(30 - utf8_length(students[i].ftName), ' ') << "|\n";
        cout << u8"-----------------------------------------------------------------------------------------------------\n|";
        for (int i3 = 0; i3 < subj_num; i3 += 2)
        {
            if (i3 == subj_num - 1)
            {
                cout << string(((grades_num[i3] - 1) * 6) + utf8_length(subj_names[i3]), '-');
                cout << '\n';

                cout << subj_names[i3] << string(((grades_num[i3] - 1) * 6) - utf8_length(subj_names[i3]) - 1, ' ') << '|' << u8"���� |\n|";
                for (int i2 = 0; i2 < grades_num[i3]; i2++)
                {
                    if (grades[i][i3][i2] < 10)
                        printf("%2.2f |", grades[i][i3][i2]);
                    else
                        printf("%2.2f|", grades[i][i3][i2]);
                }
                cout << '\n';
                break;
            }

            for (int i2 = 0; i2 < 2; i2++)
            {
                cout << subj_names[i3 + i2] << string(((grades_num[i3 + i2] - 1) * 6) - utf8_length(subj_names[i3 + i2]) - 1, ' ') << '|' << u8"���� |";
            }
            cout << "\n|";
            for (int j = 0; j < 2; j++)
            {
                for (int i2 = 0; i2 < grades_num[i3 + j]; i2++)
                {
                    if (grades[i][i3 + j][i2] < 10)
                        printf("%2.2f |", grades[i][i3 + j][i2]);
                    else
                        printf("%2.2f|", grades[i][i3 + j][i2]);
                }
            }
            cout << "\n|";
        }
    }
}


void analyzeGradesData(Student students[100], int numStudents, double grades[100][100][100], int subj_num, int grades_num[20], string subj_names[100]) {
    vector<int> maxGrades(subj_num, 0);
    vector<int> minGrades(subj_num, 100);
    vector<float> averageScores(numStudents, 0.0f);
    int failCount = 0, excellentCount = 0;

    // Analyze grades
    for (int i = 0; i < numStudents; i++)
    {
        int total = 0;
        for (int j = 0; j < subj_num; j++)
        {
            if (grades[i][j][grades_num[j] - 1] > maxGrades[j])
            {
                maxGrades[j] = grades[i][j][grades_num[j] - 1];
            }
            if (grades[i][j][grades_num[j] - 1] < minGrades[j])
            {
                minGrades[j] = grades[i][j][grades_num[j] - 1];
            }
            total += grades[i][j][grades_num[j] - 1];
        }
        averageScores[i] = total / (float)subj_num;

        // Count students who did not pass and those who scored excellently
        if (averageScores[i] < 60)
        {
            failCount++;
        }
        if (averageScores[i] >= 90)
        {
            excellentCount++;
        }
    }

    // Display analysis results
    cout << u8"�������� ���� ������� ��������:\n";
    for (int j = 0; j < subj_num; j++)
    {
        cout << u8"������� " << j + 1 << ": " << maxGrades[j] << endl;
    }
    cout << u8"\n̳���� ���� ������� ��������:\n";
    for (int j = 0; j < subj_num; j++)
    {
        cout << u8"������� " << j + 1 << ": " << minGrades[j] << endl;
    }

    cout << u8"\n������� ��� ������� ��������:\n";
    for (int i = 0; i < numStudents; i++)
    {
        cout << u8"������� " << i + 1 << " (" << students[i].lName << ' ' << students[i].fName << "): " << averageScores[i] << endl;
    }

    cout << u8"\nʳ������ ��������, �� �� ������� 60 ����: " << failCount << endl;
    cout << u8"ʳ������ ��������, �� ����� ����� 90 ����: " << excellentCount << endl;
}

bool isValidGrade(int grade)
{
    if (grade < 0 || grade > 100)
    {
        cout << "Invalid grade. Please enter a value between 0 and 100.\n";
        return false;
    }
    return true;
}

// Function to clear the input buffer
void clearBuffer() {
    cin.ignore(1000, '\n');  // Discard any remaining characters in the input buffer
}
