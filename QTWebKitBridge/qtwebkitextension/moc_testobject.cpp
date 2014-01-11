/****************************************************************************
** Meta object code from reading C++ file 'testobject.h'
**
** Created: Mon May 7 16:17:02 2012
**      by: The Qt Meta Object Compiler version 62 (Qt 4.7.3)
**
** WARNING! All changes made in this file will be lost!
*****************************************************************************/

#include "testobject.h"
#if !defined(Q_MOC_OUTPUT_REVISION)
#error "The header file 'testobject.h' doesn't include <QObject>."
#elif Q_MOC_OUTPUT_REVISION != 62
#error "This file was generated using the moc from 4.7.3. It"
#error "cannot be used with the include files from this version of Qt."
#error "(The moc has changed too much.)"
#endif

QT_BEGIN_MOC_NAMESPACE
static const uint qt_meta_data_MyApi[] = {

 // content:
       5,       // revision
       0,       // classname
       0,    0, // classinfo
       3,   14, // methods
       0,    0, // properties
       0,    0, // enums/sets
       0,    0, // constructors
       0,       // flags
       0,       // signalCount

 // slots: signature, parameters, type, tag, flags
      13,    7,    6,    6, 0x0a,
      42,   38,   34,    6, 0x0a,
      58,    6,    6,    6, 0x08,

       0        // eod
};

static const char qt_meta_stringdata_MyApi[] = {
    "MyApi\0\0param\0doSomething(QString)\0int\0"
    "a,b\0doSums(int,int)\0attachObject()\0"
};

const QMetaObject MyApi::staticMetaObject = {
    { &QObject::staticMetaObject, qt_meta_stringdata_MyApi,
      qt_meta_data_MyApi, 0 }
};

#ifdef Q_NO_DATA_RELOCATION
const QMetaObject &MyApi::getStaticMetaObject() { return staticMetaObject; }
#endif //Q_NO_DATA_RELOCATION

const QMetaObject *MyApi::metaObject() const
{
    return QObject::d_ptr->metaObject ? QObject::d_ptr->metaObject : &staticMetaObject;
}

void *MyApi::qt_metacast(const char *_clname)
{
    if (!_clname) return 0;
    if (!strcmp(_clname, qt_meta_stringdata_MyApi))
        return static_cast<void*>(const_cast< MyApi*>(this));
    return QObject::qt_metacast(_clname);
}

int MyApi::qt_metacall(QMetaObject::Call _c, int _id, void **_a)
{
    _id = QObject::qt_metacall(_c, _id, _a);
    if (_id < 0)
        return _id;
    if (_c == QMetaObject::InvokeMetaMethod) {
        switch (_id) {
        case 0: doSomething((*reinterpret_cast< const QString(*)>(_a[1]))); break;
        case 1: { int _r = doSums((*reinterpret_cast< int(*)>(_a[1])),(*reinterpret_cast< int(*)>(_a[2])));
            if (_a[0]) *reinterpret_cast< int*>(_a[0]) = _r; }  break;
        case 2: attachObject(); break;
        default: ;
        }
        _id -= 3;
    }
    return _id;
}
QT_END_MOC_NAMESPACE
