#ifndef XLSXMERGER_H
#define XLSXMERGER_H

#include <QString>
#include <QStringList>

class XlsxMerger
{
public:
    XlsxMerger();

    bool mergeFiles(const QStringList &inputPaths, const QString &outputPath);
    bool convertCsvListToXlsx(const QStringList &csvFilePaths, const QString &outputFolder);
};

#endif // XLSXMERGER_H
