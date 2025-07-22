#include "xlsxmerger.h"
#include "xlsxdocument.h"
#include "xlsxworksheet.h"
#include "conversions.h"

#include <QFileInfo>
#include <QDir>
#include <QDebug>

using namespace QXlsx;

XlsxMerger::XlsxMerger() {}

bool XlsxMerger::mergeFiles(const QStringList &inputPaths, const QString &outputPath)
{
    Document outputDoc;

    for (const QString &path : inputPaths)
    {
        Document inputDoc(path);
        if (!inputDoc.load()) {
            qWarning() << "Failed to load:" << path;
            return false;
        }

        Worksheet *inputSheet = dynamic_cast<Worksheet*>(inputDoc.currentWorksheet());
        if (!inputSheet) {
            qWarning() << "Invalid worksheet in file:" << path;
            return false;
        }

        QString sheetName = QFileInfo(path).baseName();
        outputDoc.addSheet(sheetName);
        Worksheet *outSheet = dynamic_cast<Worksheet*>(outputDoc.sheet(sheetName));
        if (!outSheet) {
            qWarning() << "Failed to create sheet:" << sheetName;
            return false;
        }

        auto dim = inputSheet->dimension();
        for (int row = dim.firstRow(); row <= dim.lastRow(); ++row) {
            for (int col = dim.firstColumn(); col <= dim.lastColumn(); ++col) {
                QVariant value = inputSheet->read(row, col);
                outSheet->write(row, col, value);
            }
        }
    }

    return outputDoc.saveAs(outputPath);
}

bool XlsxMerger::convertCsvListToXlsx(const QString &inputFolder, const QString &outputFolder)
{
    conversions conv;

    QDir inputDir(inputFolder);
    if (!inputDir.exists()) {
        qWarning() << "Input folder does not exist:" << inputFolder;
        return false;
    }

    QDir outDir(outputFolder);
    if (!outDir.exists()) {
        qWarning() << "Output folder does not exist:" << outputFolder;
        return false;
    }

    QStringList csvFiles = inputDir.entryList(QStringList() << "*.csv", QDir::Files);
    if (csvFiles.isEmpty()) {
        qWarning() << "No CSV files found in input folder:" << inputFolder;
        return false;
    }

    bool success = true;

    for (const QString &fileName : csvFiles) {
        QString fullCsvPath = inputDir.filePath(fileName);
        QString baseName = QFileInfo(fileName).completeBaseName();
        QString xlsxPath = outDir.filePath(baseName + ".xlsx");

        if (!conv.CSV_2_XLSX(fullCsvPath, xlsxPath)) {
            qWarning() << "Failed to convert:" << fullCsvPath;
            success = false;
        }
    }

    return success;
}
