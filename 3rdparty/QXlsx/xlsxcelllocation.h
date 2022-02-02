#pragma once

#include <QList>
#include <QMetaType>
#include <QObject>
#include <QSharedPointer>
#include <QString>
#include <QVector>
#include <QtGlobal>

#include "xlsxglobal.h"

QT_BEGIN_NAMESPACE_XLSX

class Cell;

class CellLocation
{
public:
	CellLocation();

	int col;
	int row;

	QSharedPointer<Cell> cell;
};

QT_END_NAMESPACE_XLSX
