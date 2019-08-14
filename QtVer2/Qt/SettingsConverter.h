#pragma once
#include <string>
#include <iostream>
#include <iomanip>
#include <map>
#include <QtCore>
#include <QJsonArray>
#include "xlsxdocument.h"

using namespace std;



class SettingsConverter
{
    size_t maxRows = 0;
    size_t maxCols = 0;
    QXlsx::Document* xlsxR = nullptr;
    QXlsx::AbstractSheet* activeSheet = nullptr;
    QString filePath = "";
    string savePath = "";
    int sheetCount = 0;
    QMap<QString, QJsonArray> valuesMap;
    QVector<QString> errorValues;
    QJsonObject valueJsonObject;
    void calculateNotEmptyRowsCount();
    void calculateNotEmptyColumnsCount();
    void addErrorToList(QString error);
    bool readXlsxFile();
    bool createJsonObject();
    bool writeJsonFile();

public:
    bool convert();
    bool openBook();
    void setActivetWorkSheet(const QString&);
    void printErrorMesseges();
    QStringList  getSheetsList();
    SettingsConverter(const QString& p);
};



