#include <QCoreApplication>
#include "Convertor.h"

int main(int argc, char *argv[])
{
    QCoreApplication a(argc, argv);
    //Convertor cnv( QString::fromStdString( "D:\\old.xlsx"));//test open
    if (argc ==2) {
        Convertor cnv(QCoreApplication::arguments().at(1));
        if (cnv.openBook()) {
            cout << "file was open" << endl;
            auto sheetList = cnv.getSheetsList();
            cnv.setActivetWorkSheet(sheetList[0]);
            cnv.convert();
            a.exec();
            a.exit(0);
        }
        else {
            cout << "cant open file" << endl;
            exit(-1);
        }
    }
    else {
        string userInput;
        cout << "input path to file" << endl;
        cout << "example: D:\\My folder\\Documents\\example.xlsx" << endl;
        getline(cin,userInput);
        Convertor cnv(QString::fromStdString(userInput));
        if (cnv.openBook()) {
            cout << "file was open" << endl;
            auto sheetList = cnv.getSheetsList();
            for (auto listElement : sheetList)
                cout << listElement.toStdString() << endl;
chooseAgain:
            cout << "choose sheet" << endl;
            int chooseNumber;
            cin >> chooseNumber;
            if (chooseNumber >= 0 && chooseNumber <= sheetList.size()) {
                cnv.setActivetWorkSheet(sheetList[chooseNumber]);
                cout << sheetList[chooseNumber].toStdString() << endl;
                cnv.convert();
            }
            else {
                cout << "incorrect list number" << endl;
                goto chooseAgain;
            }
        }
    }
}
