/********************************************************************************
** Form generated from reading UI file 'qtoperatewordwidget.ui'
**
** Created by: Qt User Interface Compiler version 5.9.8
**
** WARNING! All changes made in this file will be lost when recompiling UI file!
********************************************************************************/

#ifndef UI_QTOPERATEWORDWIDGET_H
#define UI_QTOPERATEWORDWIDGET_H

#include <QtCore/QVariant>
#include <QtWidgets/QAction>
#include <QtWidgets/QApplication>
#include <QtWidgets/QButtonGroup>
#include <QtWidgets/QHeaderView>
#include <QtWidgets/QPushButton>
#include <QtWidgets/QWidget>

QT_BEGIN_NAMESPACE

class Ui_QtOperateWordWidget
{
public:
    QPushButton *pushButton;

    void setupUi(QWidget *QtOperateWordWidget)
    {
        if (QtOperateWordWidget->objectName().isEmpty())
            QtOperateWordWidget->setObjectName(QStringLiteral("QtOperateWordWidget"));
        QtOperateWordWidget->resize(1046, 910);
        pushButton = new QPushButton(QtOperateWordWidget);
        pushButton->setObjectName(QStringLiteral("pushButton"));
        pushButton->setGeometry(QRect(50, 50, 171, 101));

        retranslateUi(QtOperateWordWidget);

        QMetaObject::connectSlotsByName(QtOperateWordWidget);
    } // setupUi

    void retranslateUi(QWidget *QtOperateWordWidget)
    {
        QtOperateWordWidget->setWindowTitle(QApplication::translate("QtOperateWordWidget", "QtOperateWordWidget", Q_NULLPTR));
        pushButton->setText(QApplication::translate("QtOperateWordWidget", "CSV\346\226\207\344\273\266\346\240\274\345\274\217\n"
"\346\265\213\350\257\225\346\212\245\345\221\212", Q_NULLPTR));
    } // retranslateUi

};

namespace Ui {
    class QtOperateWordWidget: public Ui_QtOperateWordWidget {};
} // namespace Ui

QT_END_NAMESPACE

#endif // UI_QTOPERATEWORDWIDGET_H
