// main.cpp
#include <QCoreApplication>
#include <QDebug>
#include "conversions.h"

int main(int argc, char *argv[])
{
    QCoreApplication a(argc, argv);

    QString csvPath  = QStringLiteral("C:/Users/Andrew.Wallo1/Documents/misc/testing_site/input_1.csv");
    QString xlsxPath = QStringLiteral("C:/Users/Andrew.Wallo1/Documents/misc/testing_site/output2.xlsx");

    conversions conv;

    if (!conv.CSV_2_XLSX(csvPath, xlsxPath)) {
        qCritical() << "CSV → XLSX conversion failed";
        return 1;
    }
    if (!conv.XLSX_addFilter(xlsxPath)) {
        qCritical() << "Adding AutoFilter failed";
        return 1;
    }

    qDebug() << "Done!";
    return 0;
}
