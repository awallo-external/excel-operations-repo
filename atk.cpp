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

void atk::style()
{
    Format headerFormat;
    headerFormat.setPatternBackgroundColor(QColor("#D3D3D3"));
    headerFormat.setFontBold(true);

    for (int col = 1; col <= m_lastCol; ++col)
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

QMetaType::Type atk::columnType(const QString &headerName, int maxRowsToCheck) const
{
    // 1. Find column index by header name (case-insensitive)
    int lastCol = m_sheet->dimension().lastColumn();
    int colIdx = -1;

    for (int col = 1; col <= lastCol; ++col) {
        QVariant val = m_sheet->read(1, col);
        if (val.toString().toLower() == headerName.toLower()) {
            colIdx = col;
            break;
        }
    }

    if (colIdx == -1)
        return QMetaType::Void;  // header not found

    // 2. Search downward to detect type from first non-empty cell
    int lastRow = m_sheet->dimension().lastRow();
    lastRow = std::min(lastRow, maxRowsToCheck + 1);  // +1 because header is row 1

    for (int row = 2; row <= lastRow; ++row) {
        QVariant cell = m_sheet->read(row, colIdx);
        QString text = cell.toString().trimmed();

        if (!text.isEmpty()) {
            bool ok = false;

            text.toInt(&ok);
            if (ok) return QMetaType::Int;

            text.toDouble(&ok);
            if (ok) return QMetaType::Double;

            return QMetaType::QString;
        }
    }

    // 3. No non-empty cell found
    return QMetaType::Void;
}
