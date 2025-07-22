#ifndef CONVERSIONS_H
#define CONVERSIONS_H

#include <QString>

class conversions {
public:
    conversions();

    bool CSV_2_XLSX(const QString &csvPath, const QString &xlsxPath);
    bool XLSX_addUtils(const QString &xlsxPath);
    bool XLSX_fixWidth(const QString &xlsxPath, bool isDataWidth);
};

#endif // CONVERSIONS_H
