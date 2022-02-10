#include "widget.h"
#include <QtWidgets>
#include <QApplication>
#include "test.h"
int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    Widget w;
    w.resize(800, 600);
    w.show();
   /* test t;
    t.print();*/
    return a.exec();
}
