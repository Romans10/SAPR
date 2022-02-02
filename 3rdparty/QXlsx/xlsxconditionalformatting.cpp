// xlsxconditionalformatting.cpp

#include <QDebug>
#include <QXmlStreamReader>
#include <QXmlStreamWriter>
#include <QtGlobal>

#include "xlsxcellrange.h"
#include "xlsxconditionalformatting.h"
#include "xlsxconditionalformatting_p.h"
#include "xlsxstyles_p.h"
#include "xlsxworksheet.h"

QT_BEGIN_NAMESPACE_XLSX

ConditionalFormattingPrivate::ConditionalFormattingPrivate()
{
}

ConditionalFormattingPrivate::ConditionalFormattingPrivate(const ConditionalFormattingPrivate &other)
    : QSharedData(other)
{
}

ConditionalFormattingPrivate::~ConditionalFormattingPrivate()
{
}

void ConditionalFormattingPrivate::writeCfVo(QXmlStreamWriter &writer, const XlsxCfVoData &cfvo) const
{
	writer.writeEmptyElement(QStringLiteral("cfvo"));
	QString type;
	switch (cfvo.type) {
	case ConditionalFormatting::ValueObject::Formula:
		type = QStringLiteral("formula");
		break;
	case ConditionalFormatting::ValueObject::Max:
		type = QStringLiteral("max");
		break;
	case ConditionalFormatting::ValueObject::Min:
		type = QStringLiteral("min");
		break;
	case ConditionalFormatting::ValueObject::Num:
		type = QStringLiteral("num");
		break;
	case ConditionalFormatting::ValueObject::Percent:
		type = QStringLiteral("percent");
		break;
	case ConditionalFormatting::ValueObject::Percentile:
		type = QStringLiteral("percentile");
		break;
	default:
		break;
	}
	writer.writeAttribute(QStringLiteral("type"), type);
	writer.writeAttribute(QStringLiteral("val"), cfvo.value);
	if (!cfvo.gte)
		writer.writeAttribute(QStringLiteral("gte"), QStringLiteral("0"));
}

/*!
 * \class ConditionalFormatting
 * \brief Conditional formatting for single cell or ranges
 * \inmodule QtXlsx
 *
 * The conditional formatting can be applied to a single cell or ranges of cells.
 */

/*!
    \enum ConditionalFormatting::HighlightRule

    \value HighlightRule::LessThan
    \value HighlightRule::LessThanOrEqual
    \value HighlightRule::Equal
    \value HighlightRule::NotEqual
    \value HighlightRule::GreaterThanOrEqual
    \value HighlightRule::GreaterThan
    \value HighlightRule::Between
    \value HighlightRule::NotBetween

    \value HighlightRule::ContainsText
    \value HighlightRule::NotContainsText
    \value HighlightRule::BeginsWith
    \value HighlightRule::EndsWith

    \value HighlightRule::TimePeriod

    \value HighlightRule::Duplicate
    \value HighlightRule::Unique

    \value HighlightRule::Blanks
    \value HighlightRule::NoBlanks
    \value HighlightRule::Errors
    \value HighlightRule::NoErrors

    \value HighlightRule::Top
    \value HighlightRule::TopPercent
    \value HighlightRule::Bottom
    \value HighlightRule::BottomPercent

    \value HighlightRule::AboveAverage
    \value HighlightRule::AboveOrEqualAverage
    \value HighlightRule::BelowAverage
    \value HighlightRule::BelowOrEqualAverage
    \value HighlightRule::AboveStdDev1
    \value HighlightRule::AboveStdDev2
    \value HighlightRule::AboveStdDev3
    \value HighlightRule::BelowStdDev1
    \value HighlightRule::BelowStdDev2
    \value HighlightRule::BelowStdDev3

    \value HighlightRule::Expression
*/

/*!
    \enum ConditionalFormatting::ValueObject

    \value ValueObject::Formula
    \value ValueObject::Max
    \value ValueObject::Min
    \value ValueObject::Num
    \value ValueObject::Percent
    \value ValueObject::Percentile
*/

/*!
    Construct a conditional formatting object
*/
ConditionalFormatting::ConditionalFormatting()
    : d(new ConditionalFormattingPrivate())
{
}

/*!
    Constructs a copy of \a other.
*/
ConditionalFormatting::ConditionalFormatting(const ConditionalFormatting &other)
    : d(other.d)
{
}

/*!
    Assigns \a other to this conditional formatting and returns a reference to
    this conditional formatting.
 */
ConditionalFormatting &ConditionalFormatting::operator=(const ConditionalFormatting &other)
{
	this->d = other.d;
	return *this;
}

/*!
 * Destroy the object.
 */
ConditionalFormatting::~ConditionalFormatting()
{
}

/*!
 * Add a hightlight rule with the given \a type, \a formula1, \a formula2,
 * \a format and \a stopIfTrue.
 * Return false if failed.
 */
bool ConditionalFormatting::addHighlightCellsRule(HighlightRule type, const QString &formula1, const QString &formula2, const Format &format, bool stopIfTrue)
{
	if (format.isEmpty())
		return false;

	bool skipFormula = false;

	QSharedPointer<XlsxCfRuleData> cfRule(new XlsxCfRuleData);
	if (type >= HighlightRule::LessThan && type <= HighlightRule::NotBetween) {
		cfRule->attrs[XlsxCfRuleData::Attribute::Type] = QStringLiteral("cellIs");
		QString op;
		switch (type) {
		case HighlightRule::Between:
			op = QStringLiteral("between");
			break;
		case HighlightRule::Equal:
			op = QStringLiteral("equal");
			break;
		case HighlightRule::GreaterThan:
			op = QStringLiteral("greaterThan");
			break;
		case HighlightRule::GreaterThanOrEqual:
			op = QStringLiteral("greaterThanOrEqual");
			break;
		case HighlightRule::LessThan:
			op = QStringLiteral("lessThan");
			break;
		case HighlightRule::LessThanOrEqual:
			op = QStringLiteral("lessThanOrEqual");
			break;
		case HighlightRule::NotBetween:
			op = QStringLiteral("notBetween");
			break;
		case HighlightRule::NotEqual:
			op = QStringLiteral("notEqual");
			break;
		default:
			break;
		}
		cfRule->attrs[XlsxCfRuleData::Attribute::Operator] = op;
	} else if (type >= HighlightRule::ContainsText && type <= HighlightRule::EndsWith) {
		if (type == HighlightRule::ContainsText) {
			cfRule->attrs[XlsxCfRuleData::Attribute::Type] = QStringLiteral("containsText");
			cfRule->attrs[XlsxCfRuleData::Attribute::Operator] = QStringLiteral("containsText");
			cfRule->attrs[XlsxCfRuleData::Attribute::Formula1_temp] = QStringLiteral("NOT(ISERROR(SEARCH(\"%1\",%2)))").arg(formula1);
		} else if (type == HighlightRule::NotContainsText) {
			cfRule->attrs[XlsxCfRuleData::Attribute::Type] = QStringLiteral("notContainsText");
			cfRule->attrs[XlsxCfRuleData::Attribute::Operator] = QStringLiteral("notContains");
			cfRule->attrs[XlsxCfRuleData::Attribute::Formula1_temp] = QStringLiteral("ISERROR(SEARCH(\"%2\",%1))").arg(formula1);
		} else if (type == HighlightRule::BeginsWith) {
			cfRule->attrs[XlsxCfRuleData::Attribute::Type] = QStringLiteral("beginsWith");
			cfRule->attrs[XlsxCfRuleData::Attribute::Operator] = QStringLiteral("beginsWith");
			cfRule->attrs[XlsxCfRuleData::Attribute::Formula1_temp] = QStringLiteral("LEFT(%2,LEN(\"%1\"))=\"%1\"").arg(formula1);
		} else {
			cfRule->attrs[XlsxCfRuleData::Attribute::Type] = QStringLiteral("endsWith");
			cfRule->attrs[XlsxCfRuleData::Attribute::Operator] = QStringLiteral("endsWith");
			cfRule->attrs[XlsxCfRuleData::Attribute::Formula1_temp] = QStringLiteral("RIGHT(%2,LEN(\"%1\"))=\"%1\"").arg(formula1);
		}
		cfRule->attrs[XlsxCfRuleData::Attribute::Text] = formula1;
		skipFormula = true;
	} else if (type == HighlightRule::TimePeriod) {
		cfRule->attrs[XlsxCfRuleData::Attribute::Type] = QStringLiteral("timePeriod");
		//:Todo
		return false;
	} else if (type == HighlightRule::Duplicate) {
		cfRule->attrs[XlsxCfRuleData::Attribute::Type] = QStringLiteral("duplicateValues");
	} else if (type == HighlightRule::Unique) {
		cfRule->attrs[XlsxCfRuleData::Attribute::Type] = QStringLiteral("uniqueValues");
	} else if (type == HighlightRule::Errors) {
		cfRule->attrs[XlsxCfRuleData::Attribute::Type] = QStringLiteral("containsErrors");
		cfRule->attrs[XlsxCfRuleData::Attribute::Formula1_temp] = QStringLiteral("ISERROR(%1)");
		skipFormula = true;
	} else if (type == HighlightRule::NoErrors) {
		cfRule->attrs[XlsxCfRuleData::Attribute::Type] = QStringLiteral("notContainsErrors");
		cfRule->attrs[XlsxCfRuleData::Attribute::Formula1_temp] = QStringLiteral("NOT(ISERROR(%1))");
		skipFormula = true;
	} else if (type == HighlightRule::Blanks) {
		cfRule->attrs[XlsxCfRuleData::Attribute::Type] = QStringLiteral("containsBlanks");
		cfRule->attrs[XlsxCfRuleData::Attribute::Formula1_temp] = QStringLiteral("LEN(TRIM(%1))=0");
		skipFormula = true;
	} else if (type == HighlightRule::NoBlanks) {
		cfRule->attrs[XlsxCfRuleData::Attribute::Type] = QStringLiteral("notContainsBlanks");
		cfRule->attrs[XlsxCfRuleData::Attribute::Formula1_temp] = QStringLiteral("LEN(TRIM(%1))>0");
		skipFormula = true;
	} else if (type >= HighlightRule::Top && type <= HighlightRule::BottomPercent) {
		cfRule->attrs[XlsxCfRuleData::Attribute::Type] = QStringLiteral("top10");
		if (type == HighlightRule::Bottom || type == HighlightRule::BottomPercent)
			cfRule->attrs[XlsxCfRuleData::Attribute::Bottom] = QStringLiteral("1");
		if (type == HighlightRule::TopPercent || type == HighlightRule::BottomPercent)
			cfRule->attrs[XlsxCfRuleData::Attribute::Percent] = QStringLiteral("1");
		cfRule->attrs[XlsxCfRuleData::Attribute::Rank] = !formula1.isEmpty() ? formula1 : QStringLiteral("10");
		skipFormula = true;
	} else if (type >= HighlightRule::AboveAverage && type <= HighlightRule::BelowStdDev3) {
		cfRule->attrs[XlsxCfRuleData::Attribute::Type] = QStringLiteral("aboveAverage");
		if (type >= HighlightRule::BelowAverage && type <= HighlightRule::BelowStdDev3)
			cfRule->attrs[XlsxCfRuleData::Attribute::AboveAverage] = QStringLiteral("0");
		if (type == HighlightRule::AboveOrEqualAverage || type == HighlightRule::BelowOrEqualAverage)
			cfRule->attrs[XlsxCfRuleData::Attribute::EqualAverage] = QStringLiteral("1");
		if (type == HighlightRule::AboveStdDev1 || type == HighlightRule::BelowStdDev1)
			cfRule->attrs[XlsxCfRuleData::Attribute::StdDev] = QStringLiteral("1");
		else if (type == HighlightRule::AboveStdDev2 || type == HighlightRule::BelowStdDev2)
			cfRule->attrs[XlsxCfRuleData::Attribute::StdDev] = QStringLiteral("2");
		else if (type == HighlightRule::AboveStdDev3 || type == HighlightRule::BelowStdDev3)
			cfRule->attrs[XlsxCfRuleData::Attribute::StdDev] = QStringLiteral("3");
	} else if (type == HighlightRule::Expression) {
		cfRule->attrs[XlsxCfRuleData::Attribute::Type] = QStringLiteral("expression");
	} else {
		return false;
	}

	cfRule->dxfFormat = format;
	if (stopIfTrue)
		cfRule->attrs[XlsxCfRuleData::Attribute::StopIfTrue] = true;
	if (!skipFormula) {
		if (!formula1.isEmpty())
			cfRule->attrs[XlsxCfRuleData::Attribute::Formula1] = formula1.startsWith(QLatin1String("=")) ? formula1.mid(1) : formula1;
		if (!formula2.isEmpty())
			cfRule->attrs[XlsxCfRuleData::Attribute::Formula2] = formula2.startsWith(QLatin1String("=")) ? formula2.mid(1) : formula2;
	}
	d->cfRules.append(cfRule);
	return true;
}

/*!
 * \overload
 *
 * Add a hightlight rule with the given \a type \a format and \a stopIfTrue.
 */
bool ConditionalFormatting::addHighlightCellsRule(HighlightRule type, const Format &format, bool stopIfTrue)
{
	if ((type >= HighlightRule::AboveAverage && type <= HighlightRule::BelowStdDev3) || (type >= HighlightRule::Duplicate && type <= HighlightRule::NoErrors)) {
		return addHighlightCellsRule(type, QString(), QString(), format, stopIfTrue);
	}

	return false;
}

/*!
 * \overload
 *
 * Add a hightlight rule with the given \a type, \a formula, \a format and \a stopIfTrue.
 * Return false if failed.
 */
bool ConditionalFormatting::addHighlightCellsRule(HighlightRule type, const QString &formula, const Format &format, bool stopIfTrue)
{
	if (type == HighlightRule::Between || type == HighlightRule::NotBetween)
		return false;

	return addHighlightCellsRule(type, formula, QString(), format, stopIfTrue);
}

/*!
 * Add a dataBar rule with the given \a color, \a type1, \a val1
 * , \a type2, \a val2, \a showData and \a stopIfTrue.
 * Return false if failed.
 */
bool ConditionalFormatting::addDataBarRule(
    const QColor &color, ValueObject type1, const QString &val1, ValueObject type2, const QString &val2, bool showData, bool stopIfTrue)
{
	QSharedPointer<XlsxCfRuleData> cfRule(new XlsxCfRuleData);

	cfRule->attrs[XlsxCfRuleData::Attribute::Type] = QStringLiteral("dataBar");
	cfRule->attrs[XlsxCfRuleData::Attribute::Color1] = XlsxColor(color);
	if (stopIfTrue)
		cfRule->attrs[XlsxCfRuleData::Attribute::StopIfTrue] = true;
	if (!showData)
		cfRule->attrs[XlsxCfRuleData::Attribute::HideData] = true;

	XlsxCfVoData cfvo1(type1, val1);
	XlsxCfVoData cfvo2(type2, val2);
	cfRule->attrs[XlsxCfRuleData::Attribute::Cfvo1] = QVariant::fromValue(cfvo1);
	cfRule->attrs[XlsxCfRuleData::Attribute::Cfvo2] = QVariant::fromValue(cfvo2);

	d->cfRules.append(cfRule);
	return true;
}

/*!
 * \overload
 * Add a dataBar rule with the given \a color, \a showData and \a stopIfTrue.
 */
bool ConditionalFormatting::addDataBarRule(const QColor &color, bool showData, bool stopIfTrue)
{
	return addDataBarRule(color, ValueObject::Min, QStringLiteral("0"), ValueObject::Max, QStringLiteral("0"), showData, stopIfTrue);
}

/*!
 * Add a colorScale rule with the given \a minColor, \a maxColor and \a stopIfTrue.
 * Return false if failed.
 */
bool ConditionalFormatting::add2ColorScaleRule(const QColor &minColor, const QColor &maxColor, bool stopIfTrue)
{
	ValueObject type1 = ValueObject::Min;
	ValueObject type2 = ValueObject::Max;
	QString val1 = QStringLiteral("0");
	QString val2 = QStringLiteral("0");

	QSharedPointer<XlsxCfRuleData> cfRule(new XlsxCfRuleData);

	cfRule->attrs[XlsxCfRuleData::Attribute::Type] = QStringLiteral("colorScale");
	cfRule->attrs[XlsxCfRuleData::Attribute::Color1] = XlsxColor(minColor);
	cfRule->attrs[XlsxCfRuleData::Attribute::Color2] = XlsxColor(maxColor);
	if (stopIfTrue)
		cfRule->attrs[XlsxCfRuleData::Attribute::StopIfTrue] = true;

	XlsxCfVoData cfvo1(type1, val1);
	XlsxCfVoData cfvo2(type2, val2);
	cfRule->attrs[XlsxCfRuleData::Attribute::Cfvo1] = QVariant::fromValue(cfvo1);
	cfRule->attrs[XlsxCfRuleData::Attribute::Cfvo2] = QVariant::fromValue(cfvo2);

	d->cfRules.append(cfRule);
	return true;
}

/*!
 * Add a colorScale rule with the given \a minColor, \a midColor, \a maxColor and \a stopIfTrue.
 * Return false if failed.
 */
bool ConditionalFormatting::add3ColorScaleRule(const QColor &minColor, const QColor &midColor, const QColor &maxColor, bool stopIfTrue)
{
	ValueObject type1 = ValueObject::Min;
	ValueObject type2 = ValueObject::Percent;
	ValueObject type3 = ValueObject::Max;
	QString val1 = QStringLiteral("0");
	QString val2 = QStringLiteral("50");
	QString val3 = QStringLiteral("0");

	QSharedPointer<XlsxCfRuleData> cfRule(new XlsxCfRuleData);

	cfRule->attrs[XlsxCfRuleData::Attribute::Type] = QStringLiteral("colorScale");
	cfRule->attrs[XlsxCfRuleData::Attribute::Color1] = XlsxColor(minColor);
	cfRule->attrs[XlsxCfRuleData::Attribute::Color2] = XlsxColor(midColor);
	cfRule->attrs[XlsxCfRuleData::Attribute::Color3] = XlsxColor(maxColor);

	if (stopIfTrue)
		cfRule->attrs[XlsxCfRuleData::Attribute::StopIfTrue] = true;

	XlsxCfVoData cfvo1(type1, val1);
	XlsxCfVoData cfvo2(type2, val2);
	XlsxCfVoData cfvo3(type3, val3);
	cfRule->attrs[XlsxCfRuleData::Attribute::Cfvo1] = QVariant::fromValue(cfvo1);
	cfRule->attrs[XlsxCfRuleData::Attribute::Cfvo2] = QVariant::fromValue(cfvo2);
	cfRule->attrs[XlsxCfRuleData::Attribute::Cfvo3] = QVariant::fromValue(cfvo3);

	d->cfRules.append(cfRule);
	return true;
}

/*!
    Returns the ranges on which the validation will be applied.
 */
QList<CellRange> ConditionalFormatting::ranges() const
{
	return d->ranges;
}

/*!
    Add the \a cell on which the conditional formatting will apply to.
 */
void ConditionalFormatting::addCell(const CellReference &cell)
{
	d->ranges.append(CellRange(cell, cell));
}

/*!
    \overload
    Add the cell(\a row, \a col) on which the conditional formatting will apply to.
 */
void ConditionalFormatting::addCell(int row, int col)
{
	d->ranges.append(CellRange(row, col, row, col));
}

/*!
    \overload
    Add the range(\a firstRow, \a firstCol, \a lastRow, \a lastCol) on
    which the conditional formatting will apply to.
 */
void ConditionalFormatting::addRange(int firstRow, int firstCol, int lastRow, int lastCol)
{
	d->ranges.append(CellRange(firstRow, firstCol, lastRow, lastCol));
}

/*!
    Add the \a range on which the conditional formatting will apply to.
 */
void ConditionalFormatting::addRange(const CellRange &range)
{
	d->ranges.append(range);
}

bool ConditionalFormattingPrivate::readCfRule(QXmlStreamReader &reader, XlsxCfRuleData *rule, Styles *styles)
{
	Q_ASSERT(reader.name() == QLatin1String("cfRule"));
	QXmlStreamAttributes attrs = reader.attributes();
	if (attrs.hasAttribute(QLatin1String("type")))
		rule->attrs[XlsxCfRuleData::Attribute::Type] = attrs.value(QLatin1String("type")).toString();
	if (attrs.hasAttribute(QLatin1String("dxfId"))) {
		int id = attrs.value(QLatin1String("dxfId")).toString().toInt();
		if (styles)
			rule->dxfFormat = styles->dxfFormat(id);
		else
			rule->dxfFormat.setDxfIndex(id);
	}
	rule->priority = attrs.value(QLatin1String("priority")).toString().toInt();
	if (attrs.value(QLatin1String("stopIfTrue")) == QLatin1String("1")) {
		// default is false
		rule->attrs[XlsxCfRuleData::Attribute::StopIfTrue] = QLatin1String("1");
	}
	if (attrs.value(QLatin1String("aboveAverage")) == QLatin1String("0")) {
		// default is true
		rule->attrs[XlsxCfRuleData::Attribute::AboveAverage] = QLatin1String("0");
	}
	if (attrs.value(QLatin1String("percent")) == QLatin1String("1")) {
		// default is false
		rule->attrs[XlsxCfRuleData::Attribute::Percent] = QLatin1String("1");
	}
	if (attrs.value(QLatin1String("bottom")) == QLatin1String("1")) {
		// default is false
		rule->attrs[XlsxCfRuleData::Attribute::Bottom] = QLatin1String("1");
	}
	if (attrs.hasAttribute(QLatin1String("operator")))
		rule->attrs[XlsxCfRuleData::Attribute::Operator] = attrs.value(QLatin1String("operator")).toString();

	if (attrs.hasAttribute(QLatin1String("text")))
		rule->attrs[XlsxCfRuleData::Attribute::Text] = attrs.value(QLatin1String("text")).toString();

	if (attrs.hasAttribute(QLatin1String("timePeriod")))
		rule->attrs[XlsxCfRuleData::Attribute::TimePeriod] = attrs.value(QLatin1String("timePeriod")).toString();

	if (attrs.hasAttribute(QLatin1String("rank")))
		rule->attrs[XlsxCfRuleData::Attribute::Rank] = attrs.value(QLatin1String("rank")).toString();

	if (attrs.hasAttribute(QLatin1String("stdDev")))
		rule->attrs[XlsxCfRuleData::Attribute::StdDev] = attrs.value(QLatin1String("stdDev")).toString();

	if (attrs.value(QLatin1String("equalAverage")) == QLatin1String("1")) {
		// default is false
		rule->attrs[XlsxCfRuleData::Attribute::EqualAverage] = QLatin1String("1");
	}

	while (!reader.atEnd()) {
		reader.readNextStartElement();
		if (reader.tokenType() == QXmlStreamReader::StartElement) {
			if (reader.name() == QLatin1String("formula")) {
				const QString f = reader.readElementText();
				if (!rule->attrs.contains(XlsxCfRuleData::Attribute::Formula1))
					rule->attrs[XlsxCfRuleData::Attribute::Formula1] = f;
				else if (!rule->attrs.contains(XlsxCfRuleData::Attribute::Formula2))
					rule->attrs[XlsxCfRuleData::Attribute::Formula2] = f;
				else if (!rule->attrs.contains(XlsxCfRuleData::Attribute::Formula3))
					rule->attrs[XlsxCfRuleData::Attribute::Formula3] = f;
			} else if (reader.name() == QLatin1String("dataBar")) {
				readCfDataBar(reader, rule);
			} else if (reader.name() == QLatin1String("colorScale")) {
				readCfColorScale(reader, rule);
			}
		}
		if (reader.tokenType() == QXmlStreamReader::EndElement && reader.name() == QStringLiteral("conditionalFormatting")) {
			break;
		}
	}
	return true;
}

bool ConditionalFormattingPrivate::readCfDataBar(QXmlStreamReader &reader, XlsxCfRuleData *rule)
{
	Q_ASSERT(reader.name() == QLatin1String("dataBar"));
	QXmlStreamAttributes attrs = reader.attributes();
	if (attrs.value(QLatin1String("showValue")) == QLatin1String("0"))
		rule->attrs[XlsxCfRuleData::Attribute::HideData] = QStringLiteral("1");

	while (!reader.atEnd()) {
		reader.readNextStartElement();
		if (reader.tokenType() == QXmlStreamReader::StartElement) {
			if (reader.name() == QLatin1String("cfvo")) {
				XlsxCfVoData data;
				readCfVo(reader, data);
				if (!rule->attrs.contains(XlsxCfRuleData::Attribute::Cfvo1))
					rule->attrs[XlsxCfRuleData::Attribute::Cfvo1] = QVariant::fromValue(data);
				else
					rule->attrs[XlsxCfRuleData::Attribute::Cfvo2] = QVariant::fromValue(data);
			} else if (reader.name() == QLatin1String("color")) {
				XlsxColor color;
				color.loadFromXml(reader);
				rule->attrs[XlsxCfRuleData::Attribute::Color1] = color;
			}
		}
		if (reader.tokenType() == QXmlStreamReader::EndElement && reader.name() == QStringLiteral("dataBar")) {
			break;
		}
	}

	return true;
}

bool ConditionalFormattingPrivate::readCfColorScale(QXmlStreamReader &reader, XlsxCfRuleData *rule)
{
	Q_ASSERT(reader.name() == QLatin1String("colorScale"));

	while (!reader.atEnd()) {
		reader.readNextStartElement();
		if (reader.tokenType() == QXmlStreamReader::StartElement) {
			if (reader.name() == QLatin1String("cfvo")) {
				XlsxCfVoData data;
				readCfVo(reader, data);
				if (!rule->attrs.contains(XlsxCfRuleData::Attribute::Cfvo1))
					rule->attrs[XlsxCfRuleData::Attribute::Cfvo1] = QVariant::fromValue(data);
				else if (!rule->attrs.contains(XlsxCfRuleData::Attribute::Cfvo2))
					rule->attrs[XlsxCfRuleData::Attribute::Cfvo2] = QVariant::fromValue(data);
				else
					rule->attrs[XlsxCfRuleData::Attribute::Cfvo2] = QVariant::fromValue(data);
			} else if (reader.name() == QLatin1String("color")) {
				XlsxColor color;
				color.loadFromXml(reader);
				if (!rule->attrs.contains(XlsxCfRuleData::Attribute::Color1))
					rule->attrs[XlsxCfRuleData::Attribute::Color1] = color;
				else if (!rule->attrs.contains(XlsxCfRuleData::Attribute::Color2))
					rule->attrs[XlsxCfRuleData::Attribute::Color2] = color;
				else
					rule->attrs[XlsxCfRuleData::Attribute::Color3] = color;
			}
		}
		if (reader.tokenType() == QXmlStreamReader::EndElement && reader.name() == QStringLiteral("colorScale")) {
			break;
		}
	}

	return true;
}

bool ConditionalFormattingPrivate::readCfVo(QXmlStreamReader &reader, XlsxCfVoData &cfvo)
{
	Q_ASSERT(reader.name() == QStringLiteral("cfvo"));

	QXmlStreamAttributes attrs = reader.attributes();

	QString type = attrs.value(QLatin1String("type")).toString();
	ConditionalFormatting::ValueObject t;
	if (type == QLatin1String("formula"))
		t = ConditionalFormatting::ValueObject::Formula;
	else if (type == QLatin1String("max"))
		t = ConditionalFormatting::ValueObject::Max;
	else if (type == QLatin1String("min"))
		t = ConditionalFormatting::ValueObject::Min;
	else if (type == QLatin1String("num"))
		t = ConditionalFormatting::ValueObject::Num;
	else if (type == QLatin1String("percent"))
		t = ConditionalFormatting::ValueObject::Percent;
	else  // if (type == QLatin1String("percentile"))
		t = ConditionalFormatting::ValueObject::Percentile;

	cfvo.type = t;
	cfvo.value = attrs.value(QLatin1String("val")).toString();
	if (attrs.value(QLatin1String("gte")) == QLatin1String("0")) {
		// default is true
		cfvo.gte = false;
	}
	return true;
}

bool ConditionalFormatting::loadFromXml(QXmlStreamReader &reader, Styles *styles)
{
	Q_ASSERT(reader.name() == QStringLiteral("conditionalFormatting"));

	d->ranges.clear();
	d->cfRules.clear();
	QXmlStreamAttributes attrs = reader.attributes();
	const QString sqref = attrs.value(QLatin1String("sqref")).toString();
	const auto sqrefParts = sqref.split(QLatin1Char(' '));
	for (const QString &range : sqrefParts) {
		this->addRange(CellRange {range});
	}

	while (!reader.atEnd()) {
		reader.readNextStartElement();
		if (reader.tokenType() == QXmlStreamReader::StartElement) {
			if (reader.name() == QLatin1String("cfRule")) {
				QSharedPointer<XlsxCfRuleData> cfRule(new XlsxCfRuleData);
				d->readCfRule(reader, cfRule.data(), styles);
				d->cfRules.append(cfRule);
			}
		}
		if (reader.tokenType() == QXmlStreamReader::EndElement && reader.name() == QStringLiteral("conditionalFormatting")) {
			break;
		}
	}

	return true;
}

bool ConditionalFormatting::saveToXml(QXmlStreamWriter &writer) const
{
	writer.writeStartElement(QStringLiteral("conditionalFormatting"));
	QStringList sqref;
	const auto rangeList = ranges();
	for (const CellRange &range : rangeList) {
		sqref.append(range.toString());
	}
	writer.writeAttribute(QStringLiteral("sqref"), sqref.join(QLatin1String(" ")));

	for (int i = 0; i < d->cfRules.size(); ++i) {
		const QSharedPointer<XlsxCfRuleData> &rule = d->cfRules[i];
		writer.writeStartElement(QStringLiteral("cfRule"));
		writer.writeAttribute(QStringLiteral("type"), rule->attrs[XlsxCfRuleData::Attribute::Type].toString());
		if (rule->dxfFormat.dxfIndexValid())
			writer.writeAttribute(QStringLiteral("dxfId"), QString::number(rule->dxfFormat.dxfIndex()));
		writer.writeAttribute(QStringLiteral("priority"), QString::number(rule->priority));

		auto it = rule->attrs.constFind(XlsxCfRuleData::Attribute::StopIfTrue);
		if (it != rule->attrs.constEnd())
			writer.writeAttribute(QStringLiteral("stopIfTrue"), it.value().toString());

		it = rule->attrs.constFind(XlsxCfRuleData::Attribute::AboveAverage);
		if (it != rule->attrs.constEnd())
			writer.writeAttribute(QStringLiteral("aboveAverage"), it.value().toString());

		it = rule->attrs.constFind(XlsxCfRuleData::Attribute::Percent);
		if (it != rule->attrs.constEnd())
			writer.writeAttribute(QStringLiteral("percent"), it.value().toString());

		it = rule->attrs.constFind(XlsxCfRuleData::Attribute::Bottom);
		if (it != rule->attrs.constEnd())
			writer.writeAttribute(QStringLiteral("bottom"), it.value().toString());

		it = rule->attrs.constFind(XlsxCfRuleData::Attribute::Operator);
		if (it != rule->attrs.constEnd())
			writer.writeAttribute(QStringLiteral("operator"), it.value().toString());

		it = rule->attrs.constFind(XlsxCfRuleData::Attribute::Text);
		if (it != rule->attrs.constEnd())
			writer.writeAttribute(QStringLiteral("text"), it.value().toString());

		it = rule->attrs.constFind(XlsxCfRuleData::Attribute::TimePeriod);
		if (it != rule->attrs.constEnd())
			writer.writeAttribute(QStringLiteral("timePeriod"), it.value().toString());

		it = rule->attrs.constFind(XlsxCfRuleData::Attribute::Rank);
		if (it != rule->attrs.constEnd())
			writer.writeAttribute(QStringLiteral("rank"), it.value().toString());

		it = rule->attrs.constFind(XlsxCfRuleData::Attribute::StdDev);
		if (it != rule->attrs.constEnd())
			writer.writeAttribute(QStringLiteral("stdDev"), it.value().toString());

		it = rule->attrs.constFind(XlsxCfRuleData::Attribute::EqualAverage);
		if (it != rule->attrs.constEnd())
			writer.writeAttribute(QStringLiteral("equalAverage"), it.value().toString());

		if (rule->attrs[XlsxCfRuleData::Attribute::Type] == QLatin1String("dataBar")) {
			writer.writeStartElement(QStringLiteral("dataBar"));
			if (rule->attrs.contains(XlsxCfRuleData::Attribute::HideData))
				writer.writeAttribute(QStringLiteral("showValue"), QStringLiteral("0"));
			d->writeCfVo(writer, rule->attrs[XlsxCfRuleData::Attribute::Cfvo1].value<XlsxCfVoData>());
			d->writeCfVo(writer, rule->attrs[XlsxCfRuleData::Attribute::Cfvo2].value<XlsxCfVoData>());
			rule->attrs[XlsxCfRuleData::Attribute::Color1].value<XlsxColor>().saveToXml(writer);
			writer.writeEndElement();  // dataBar
		} else if (rule->attrs[XlsxCfRuleData::Attribute::Type] == QLatin1String("colorScale")) {
			writer.writeStartElement(QStringLiteral("colorScale"));
			d->writeCfVo(writer, rule->attrs[XlsxCfRuleData::Attribute::Cfvo1].value<XlsxCfVoData>());
			d->writeCfVo(writer, rule->attrs[XlsxCfRuleData::Attribute::Cfvo2].value<XlsxCfVoData>());

			it = rule->attrs.constFind(XlsxCfRuleData::Attribute::Cfvo3);
			if (it != rule->attrs.constEnd())
				d->writeCfVo(writer, it.value().value<XlsxCfVoData>());

			rule->attrs[XlsxCfRuleData::Attribute::Color1].value<XlsxColor>().saveToXml(writer);
			rule->attrs[XlsxCfRuleData::Attribute::Color2].value<XlsxColor>().saveToXml(writer);

			it = rule->attrs.constFind(XlsxCfRuleData::Attribute::Color3);
			if (it != rule->attrs.constEnd())
				it.value().value<XlsxColor>().saveToXml(writer);

			writer.writeEndElement();  // colorScale
		}

		it = rule->attrs.constFind(XlsxCfRuleData::Attribute::Formula1_temp);
		if (it != rule->attrs.constEnd()) {
			QString str = (ranges().begin())->toString();
			QString startCell = *(str.split(QLatin1Char(':')).begin());
			writer.writeTextElement(QStringLiteral("formula"), it.value().toString().arg(startCell));
		} else if ((it = rule->attrs.constFind(XlsxCfRuleData::Attribute::Formula1)) != rule->attrs.constEnd()) {
			writer.writeTextElement(QStringLiteral("formula"), it.value().toString());
		}

		it = rule->attrs.constFind(XlsxCfRuleData::Attribute::Formula2);
		if (it != rule->attrs.constEnd())
			writer.writeTextElement(QStringLiteral("formula"), it.value().toString());

		it = rule->attrs.constFind(XlsxCfRuleData::Attribute::Formula3);
		if (it != rule->attrs.constEnd())
			writer.writeTextElement(QStringLiteral("formula"), it.value().toString());

		writer.writeEndElement();  // cfRule
	}

	writer.writeEndElement();  // conditionalFormatting
	return true;
}

QT_END_NAMESPACE_XLSX
