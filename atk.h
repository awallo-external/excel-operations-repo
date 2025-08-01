// atk.h
// Requires Qt and QXlsx for functionality

#ifndef ATK_H
#define ATK_H

#include <QString>
#include "xlsxmerger.h"
#include "xlsxdocument.h"
#include "xlsxworksheet.h"
#include "conversions.h"

#include <QFileInfo>
#include <QDir>
#include <QDebug>


using namespace QXlsx;

class atk
{
public:
    atk(QXlsx::Worksheet *sheet);  // (otherwise my functions can't find the sheet)

    void addFilter(int lastCol);
    void freezeTopRow(bool active);
    void style();
    int computeLastCol();
    QMetaType::Type columnType(const QString &headerName, int maxRowsToCheck = 100) const;


private:
    QXlsx::Worksheet *m_sheet;  // Store the sheet
    int m_lastCol;              // final column that has data

};

#endif // ATK_H
