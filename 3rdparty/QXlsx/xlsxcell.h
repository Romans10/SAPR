#pragma once

#include <cstdio>

#include <QDate>
#include <QDateTime>
#include <QObject>
#include <QString>
#include <QTime>
#include <QVariant>
#include <QtGlobal>

#include "xlsxformat.h"
#include "xlsxglobal.h"

QT_BEGIN_NAMESPACE_XLSX

class Worksheet;
class Format;
class CellFormula;
class CellPrivate;
class WorksheetPrivate;

class Cell
{
	Q_DECLARE_PRIVATE(Cell)

public:
	enum class Type  // See ECMA 376, 18.18.11. ST_CellType (Cell Type) for more information.
	{
		Boolean,
		Date,
		InlineString,
		Number,
		SharedString,
		String,
		Custom,  // custom or un-defined cell type
		Error,
	};

	Cell(const QVariant &data = QVariant(), Type type = Type::Number, const Format &format = Format(), Worksheet *parent = NULL, qint32 styleIndex = (-1));

	Cell(const Cell *const cell);
	~Cell();

	CellPrivate *const d_ptr;  // See D-pointer and Q-pointer of Qt, for more information.

	Type type() const;
	QVariant value() const;
	QVariant readValue() const;
	Format format() const;

	bool hasFormula() const;
	CellFormula formula() const;

	bool isDateTime() const;
	QVariant dateTime() const;  // QDateTime, QDate, QTime

	bool isRichString() const;

	qint32 styleNumber() const;

private:
	friend class Worksheet;
	friend class WorksheetPrivate;
};

QT_END_NAMESPACE_XLSX
