#include "Convertor.h"

int main(int argc, char *argv[])
{
    QCoreApplication a(argc, argv);
    QTextStream qtin(stdin);
    QTextStream qtout(stdout);
    QString fileQStrPath;
    string fileStdStrPath;
    bool printDataFlag = true;
    bool printNameFlag = true;
    int userChoose1;
    int userChoose2;
    qtout << "welcome to xlsx_to_json converter" << endl;
    qtout << "enter  path to file:" << endl;
    qtout << "example: 'D:\\folder1\\folder2\\file.xlsx'" << endl;
    getline(cin,fileStdStrPath);
    fileQStrPath = QFileInfo (QString::fromStdString(fileStdStrPath)).absoluteFilePath();
    qtout << " want you print names?" << endl;
    qtout << "press 0 if not? else press any key" << endl;
    qtin >> userChoose1;
    printNameFlag = userChoose1;
    qtout << " want you  also print data to console?" << endl;
    qtout << "press 0 if not? else press any key" << endl;
    qtin >> userChoose2;
    printDataFlag = userChoose2;
    Convertor cnv(fileQStrPath,printDataFlag,printNameFlag);
    if (cnv.openBook()) {
        cout << "file is open" << endl;
        cnv.convert();
        cout << endl;
    }
    else {
        cout << " Cant open File" << endl;
        exit(-1);
    }
    return a.exec();
}

