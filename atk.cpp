/* atk.cpp */

// andy tool kit

#include "atk.h"



using namespace QXlsx;

atk::atk(QXlsx::Worksheet *sheet)
    : m_sheet(sheet)
{
    m_lastCol = computeLastCol();
}

void atk::addFilter(int lastCol)
{
    m_sheet->setAutoFilter(QXlsx::CellRange(1,1,1,lastCol));
}

void atk::freezeTopRow(bool active)
{
    m_sheet->setFreezeTopRow(active);
}

void atk::style(int lastCol)
{
    Format headerFormat;
    headerFormat.setPatternBackgroundColor(QColor("#D3D3D3"));
    headerFormat.setFontBold(true);

    for (int col = 1; col <= lastCol; ++col)
    {
        QVariant val = m_sheet->read(1,col);
        m_sheet->write(1,col, val, headerFormat);
    }

}

int atk::computeLastCol()
{
    int col = 1;
    while (!m_sheet->read(1, col).toString().isEmpty())
        ++col;
    return col - 1;
}
