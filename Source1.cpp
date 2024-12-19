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

void create_Sample(string subj_names[100], int subj_num, int grades_num[20])
{
    // ��������� �����
    xlnt::workbook wb;

    // ��������� ������
    wb.create_sheet();

    // ����� ������ � ������
    auto sheet = wb.active_sheet();

    int i2 = 0;
    for (int i = 8; i2 < subj_num; i += (grades_num[i2 - 1] + 1))
    {
        xlnt::cell_reference cell_ref(i, 1);
        sheet.cell(cell_ref).value(subj_names[i2++]);
        cell_ref.column_index(i + grades_num[i2 - 1]);
        cell_ref.row(3);
        sheet.cell(cell_ref).value(u8"����");
    }

    xlnt::cell_reference cell_ref(1, 3);
    sheet.cell(cell_ref).value(u8"�");
    cell_ref.column_index(2);
    cell_ref.row(3);
    sheet.cell(cell_ref).value(u8"�������");
    cell_ref.column_index(3);
    cell_ref.row(3);
    sheet.cell(cell_ref).value(u8"��'�");
    cell_ref.column_index(4);
    cell_ref.row(3);
    sheet.cell(cell_ref).value(u8"�� �������");
    cell_ref.column_index(5);
    cell_ref.row(3);
    sheet.cell(cell_ref).value(u8"�������");
    cell_ref.column_index(6);
    cell_ref.row(3);
    sheet.cell(cell_ref).value(u8"̳���");

    // ���������� �����
    wb.save(u8"Sample.xlsx");

    std::cout << u8"������ ��������: Sample.xlsx" << std::endl;
}

void load_xlsx(xlnt::workbook wb, string subj_names[100], int subj_num, int grades_num[20], Student students[100], int numStudents, double grades[100][100][100])
{
    // �������� �������� �����
    auto sheet = wb.sheet_by_title(u8"Sheet1");

    for (int i = 4; i <= numStudents + 4; i++)
    {
        xlnt::cell_reference cell_ref(2, i);
        students[i - 4].lName = sheet.cell(cell_ref).to_string();
        cell_ref.column_index(3);
        students[i - 4].fName = sheet.cell(cell_ref).to_string();
        cell_ref.column_index(4);
        students[i - 4].ftName = sheet.cell(cell_ref).to_string();
        cell_ref.column_index(5);
        students[i - 4].phone = sheet.cell(cell_ref).to_string();
        cell_ref.column_index(6);
        students[i - 4].city = sheet.cell(cell_ref).to_string();

        int i3 = 0;
        for (int s = 8; i3 < subj_num; s += (grades_num[i3 - 1] + 1))
        {
            for (int g = s; g - s < grades_num[i3]; g++)
            {
                xlnt::cell_reference cell_ref(g, i);
                if (sheet.cell(cell_ref).has_value() || sheet.cell(cell_ref).has_formula())
                {
                    grades[i - 4][i3][g - s] = sheet.cell(cell_ref).value<double>();
                }
                else
                {
                    grades[i - 4][i3][g - s] = 0;
                }
            }
            i3++;
        }
    }
    int i3 = 0;
    for (int s = 8; i3 < subj_num; s += (grades_num[i3 - 1] + 1))
    {
        xlnt::cell_reference cell_ref(s, 1);
        subj_names[i3++] = sheet.cell(cell_ref).to_string();
    }
}

void save_xlsx(xlnt::workbook wb, xlnt::workbook new_wb, string subj_names[100], int subj_num, int grades_num[20], Student students[100], int numStudents, double grades[100][100][100])
{
    // ����������� �������� ����
    xlnt::worksheet source_ws = wb.sheet_by_title("Sheet1");

    // �������� ����� ���� � ������ ������� ����
    xlnt::worksheet new_ws = new_wb.sheet_by_title("Sheet1");

    // ��������� ������ ��������
    for (const auto& col : source_ws.columns()) {
        int col_index = col.front().reference().column_index();
        new_ws.column_properties(col_index).width = source_ws.column_properties(col_index).width;
    }

    // ��������� ������ �����
    for (const auto& row : source_ws.rows()) {
        int row_index = row.front().reference().row();
        new_ws.row_properties(row_index).height = source_ws.row_properties(row_index).height;
    }

    // ��������� ��� �� ���� �������
    for (auto row : source_ws.rows()) {
        for (auto cell : row) {
            auto ref = cell.reference();

            new_ws.cell(ref).value(cell.to_string());

            new_ws.cell(ref).fill(cell.fill());

            new_ws.cell(ref).font(cell.font());

            new_ws.cell(ref).border(cell.border());
        }
    }


    auto sheet = new_wb.sheet_by_title(u8"Sheet1");

    for (int i = 4; i < numStudents + 4; i++)
    {
        xlnt::cell_reference cell_ref(1, i);
        sheet.cell(cell_ref).value(i - 3);
        cell_ref.column_index(2);
        sheet.cell(cell_ref).value(students[i - 4].lName);
        cell_ref.column_index(3);
        sheet.cell(cell_ref).value(students[i - 4].fName);
        cell_ref.column_index(4);
        sheet.cell(cell_ref).value(students[i - 4].ftName);
        cell_ref.column_index(5);
        sheet.cell(cell_ref).value(students[i - 4].phone);
        cell_ref.column_index(6);
        sheet.cell(cell_ref).value(students[i - 4].city);

        int i3 = 0;
        for (int s = 8; i3 < subj_num; s += (grades_num[i3 - 1] + 1))
        {
            for (int g = s; g - s < grades_num[i3]; g++)
            {
                xlnt::cell_reference cell_ref(g, i);
                sheet.cell(cell_ref).value(grades[i - 4][i3][g - s]);
            }
            i3++;
        }
    }
    int i3 = 0;
    for (int s = 8; i3 < subj_num; s += (grades_num[i3 - 1] + 1))
    {
        xlnt::cell_reference cell_ref(s, 1);
        sheet.cell(cell_ref).value(subj_names[i3++]);
    }

    new_wb.save(u8"saved.xlsx");
}



/*void save_xlsx(xlnt::workbook wb, string subj_names[100], int subj_num, int grades_num[20], Student students[100], int numStudents, double grades[100][100][100])
{
    // �������� �������� �����
    auto sheet = wb.sheet_by_title(u8"Sheet1");

    for (int i = 4; i < numStudents + 4; i++)
    {
        xlnt::cell_reference cell_ref(1, i);
        sheet.cell(cell_ref).value(i - 3);
        cell_ref.column_index(2);
        sheet.cell(cell_ref).value(students[i - 4].lName);
        cell_ref.column_index(3);
        sheet.cell(cell_ref).value(students[i - 4].fName);
        cell_ref.column_index(4);
        sheet.cell(cell_ref).value(students[i - 4].ftName);
        cell_ref.column_index(5);
        sheet.cell(cell_ref).value(students[i - 4].phone);
        cell_ref.column_index(6);
        sheet.cell(cell_ref).value(students[i - 4].city);

        int i3 = 0;
        for (int s = 8; i3 < subj_num; s += (grades_num[i3 - 1] + 1))
        {
            for (int g = s; g - s < grades_num[i3]; g++)
            {
                xlnt::cell_reference cell_ref(g, i);
                sheet.cell(cell_ref).value(grades[i - 4][i3][g - s]);
            }
            i3++;
        }
    }
    int i3 = 0;
    for (int s = 8; i3 < subj_num; s += (grades_num[i3 - 1] + 1))
    {
        xlnt::cell_reference cell_ref(s, 1);
        sheet.cell(cell_ref).value(subj_names[i3++]);
    }

    wb.save(u8"saved.xlsx");
}

void copyProp(xlnt::workbook source_wb, xlnt::workbook new_wb)
{
    // ����������� �������� ����
    xlnt::worksheet source_ws = source_wb.sheet_by_title("Sheet1");

    // �������� ����� ���� � ������ ������� ����
    xlnt::worksheet new_ws = new_wb.sheet_by_title("Sheet1");

    // ��������� ������ ��������
    for (const auto& col : source_ws.columns()) {
        int col_index = col.front().reference().column_index();
        new_ws.column_properties(col_index).width = source_ws.column_properties(col_index).width;
    }

    // ��������� ������ �����
    for (const auto& row : source_ws.rows()) {
        int row_index = row.front().reference().row();
        new_ws.row_properties(row_index).height = source_ws.row_properties(row_index).height;
    }

    // ��������� ��� �� ���� �������
    for (auto row : source_ws.rows()) {
        for (auto cell : row) {
            auto ref = cell.reference();
            // ��������� ��������
            new_ws.cell(ref).value(cell.to_string());

            // ��������� ������� (���� ����)
            new_ws.cell(ref).fill(cell.fill());

            // ��������� ����� (�� �����������)
            new_ws.cell(ref).font(cell.font());
        }
    }
    new_wb.save("zxc.xlsx");
    return;
}
*/