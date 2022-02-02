#pragma once

#include <QColor>
#include <QList>
#include <QSharedDataPointer>
#include <QString>
#include <QXmlStreamReader>
#include <QXmlStreamWriter>
#include <QtGlobal>

#include "xlsxcellrange.h"
#include "xlsxcellreference.h"
#include "xlsxglobal.h"

class ConditionalFormattingTest;

QT_BEGIN_NAMESPACE_XLSX

class Format;
class Worksheet;
class Styles;
class ConditionalFormattingPrivate;

class ConditionalFormatting
{
public:
	enum class HighlightRule
	{
		LessThan,
		LessThanOrEqual,
		Equal,
		NotEqual,
		GreaterThanOrEqual,
		GreaterThan,
		Between,
		NotBetween,

		ContainsText,
		NotContainsText,
		BeginsWith,
		EndsWith,

		TimePeriod,

		Duplicate,
		Unique,
		Blanks,
		NoBlanks,
		Errors,
		NoErrors,

		Top,
		TopPercent,
		Bottom,
		BottomPercent,

		AboveAverage,
		AboveOrEqualAverage,
		AboveStdDev1,
		AboveStdDev2,
		AboveStdDev3,
		BelowAverage,
		BelowOrEqualAverage,
		BelowStdDev1,
		BelowStdDev2,
		BelowStdDev3,

		Expression
	};

	enum class ValueObject
	{
		Formula,
		Max,
		Min,
		Num,
		Percent,
		Percentile
	};

public:
	ConditionalFormatting();
	ConditionalFormatting(const ConditionalFormatting &other);
	~ConditionalFormatting();

	bool addHighlightCellsRule(HighlightRule type, const Format &format, bool stopIfTrue = false);
	bool addHighlightCellsRule(HighlightRule type, const QString &formula1, const Format &format, bool stopIfTrue = false);
	bool addHighlightCellsRule(HighlightRule type, const QString &formula1, const QString &formula2, const Format &format, bool stopIfTrue = false);
	bool addDataBarRule(const QColor &color, bool showData = true, bool stopIfTrue = false);
	bool addDataBarRule(
	    const QColor &color, ValueObject type1, const QString &val1, ValueObject type2, const QString &val2, bool showData = true, bool stopIfTrue = false);
	bool add2ColorScaleRule(const QColor &minColor, const QColor &maxColor, bool stopIfTrue = false);
	bool add3ColorScaleRule(const QColor &minColor, const QColor &midColor, const QColor &maxColor, bool stopIfTrue = false);

	QList<CellRange> ranges() const;

	void addCell(const CellReference &cell);
	void addCell(int row, int col);
	void addRange(int firstRow, int firstCol, int lastRow, int lastCol);
	void addRange(const CellRange &range);

	// needed by QSharedDataPointer!!
	ConditionalFormatting &operator=(const ConditionalFormatting &other);

private:
	friend class Worksheet;
	friend class ::ConditionalFormattingTest;

	bool saveToXml(QXmlStreamWriter &writer) const;
	bool loadFromXml(QXmlStreamReader &reader, Styles *styles = NULL);

	QSharedDataPointer<ConditionalFormattingPrivate> d;
};

QT_END_NAMESPACE_XLSX
