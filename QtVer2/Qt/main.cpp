#include <QCoreApplication>
#include "SettingsConverter.h"

int main(int argc, char *argv[])
{
    QCoreApplication a(argc, argv);
    if (argc ==2) {
        SettingsConverter cnv(QCoreApplication::arguments().at(1));
        if (cnv.openBook())
        {
            cout << "file was open" << endl;
            auto sheetList = cnv.getSheetsList();
            cnv.setActivetWorkSheet(sheetList[0]);
            if (cnv.convert()) {
                a.exec();
                a.exit(0);
            }
            else {
                cnv.printErrorMesseges();
                exit(-1);
            }
        }  //argc ==2  // if(cnv.openBook())
        else {
            cnv.printErrorMesseges();
            exit(-1);
        }
    } // if(argc == 2)
    else  {
        string userInput;
        cout << "input path to file" << endl;
        cout << "example: D:\\My folder\\Documents\\example.xlsx" << endl;
        getline(cin,userInput);
        SettingsConverter cnv(QString::fromStdString(userInput));
        if (cnv.openBook()) {
            cout << "file was open" << endl;
            auto sheetList = cnv.getSheetsList();
            for (auto listElement : sheetList)
            {
                cout << listElement.toStdString() << endl;
            }
            int userInputTry = 0;
            while (userInputTry == 0) {
                cout << "choose sheet" << endl;
                int chooseNumber;
                cin >> chooseNumber;
                if (chooseNumber >= 0 && chooseNumber <= sheetList.size()) {
                    userInputTry++;
                    cnv.setActivetWorkSheet(sheetList[chooseNumber]);
                    cout << sheetList[chooseNumber].toStdString() << endl;
                    if (cnv.convert()) {
                        a.exit(0);

                    }
                    else {
                        cnv.printErrorMesseges();
                        a.exit(-1);
                    }
                }
                else {
                    cout << "incorrect list number" << endl;
                }
            }
        }  //if(argc != 2) // if(cnv.openBook())
        else {
            cnv.printErrorMesseges();
            exit(-1);
        }
    }
}
