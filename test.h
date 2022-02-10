#ifndef TEST_H
#define TEST_H

#include <QObject>
#include <QAxObject>
class test : public QObject
{
    Q_OBJECT
public:
    explicit test(QObject *parent = nullptr);
    void print();
signals:

};

#endif // TEST_H
