#pragma once

#include <QList>
#include <QObject>
#include <QSharedPointer>
#include <QtGlobal>

#include "xlsxcell.h"
#include "xlsxcellformula.h"
#include "xlsxcellrange.h"
#include "xlsxglobal.h"
#include "xlsxrichstring.h"

QT_BEGIN_NAMESPACE_XLSX

class CellPrivate
{
	Q_DECLARE_PUBLIC(Cell)

public:
	CellPrivate(Cell *p);
	CellPrivate(const CellPrivate *const cp);

	Worksheet *parent;
	Cell *q_ptr;

	Cell::Type cellType;
	QVariant value;

	CellFormula formula;
	Format format;

	RichString richString;

	qint32 styleNumber;
};

QT_END_NAMESPACE_XLSX
