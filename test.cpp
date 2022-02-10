#include "test.h"

test::test(QObject *parent) : QObject(parent)
{

}
void test::print(){
    QAxObject * word = new QAxObject("Word.Application");
        word->setProperty("DisplayAlerts", "0");
        word->querySubObject("Documents")->querySubObject("Open(QVariant)", "d:\\txt\\1.xml");
        // Печатаем
        word->querySubObject("ActiveDocument")->dynamicCall("PrintOut()");
        // Закрываем
        word->querySubObject("ActiveDocument")->dynamicCall("Close()");
        word->dynamicCall("Quit()");
}
