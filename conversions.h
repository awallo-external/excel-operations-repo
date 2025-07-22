#ifndef CONVERSIONS_H
#define CONVERSIONS_H

#include <QString>

class conversions {
public:
    conversions();

    bool CSV_2_XLSX(const QString &csvPath, const QString &xlsxPath);
    bool XLSX_addFilter(const QString &xlsxPath);
};

#endif // CONVERSIONS_H
