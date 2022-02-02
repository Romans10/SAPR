#pragma once

#include <QMap>
#include <QSharedData>
#include <QSharedPointer>
#include <QtGlobal>

#include "xlsxcolor_p.h"
#include "xlsxconditionalformatting.h"
#include "xlsxformat.h"

QT_BEGIN_NAMESPACE_XLSX

class XlsxCfVoData
{
public:
	XlsxCfVoData()
	    : gte(true)
	{
	}

	XlsxCfVoData(ConditionalFormatting::ValueObject type, const QString &value, bool gte = true)
	    : type(type)
	    , value(value)
	    , gte(gte)
	{
	}

	ConditionalFormatting::ValueObject type;
	QString value;
	bool gte;
};

class XlsxCfRuleData
{
public:
	enum class Attribute
	{
		Type,
		DxfId,
		// Priority,
		StopIfTrue,
		AboveAverage,
		Percent,
		Bottom,
		Operator,
		Text,
		TimePeriod,
		Rank,
		StdDev,
		EqualAverage,

		DxfFormat,
		Formula1,
		Formula2,
		Formula3,
		Formula1_temp,

		Color1,
		Color2,
		Color3,

		Cfvo1,
		Cfvo2,
		Cfvo3,

		HideData
	};

	XlsxCfRuleData()
	    : priority(1)
	{
	}

	int priority;
	Format dxfFormat;
	QMap<Attribute, QVariant> attrs;
};

class ConditionalFormattingPrivate : public QSharedData
{
public:
	ConditionalFormattingPrivate();
	ConditionalFormattingPrivate(const ConditionalFormattingPrivate &other);
	~ConditionalFormattingPrivate();

	void writeCfVo(QXmlStreamWriter &writer, const XlsxCfVoData &cfvo) const;
	bool readCfVo(QXmlStreamReader &reader, XlsxCfVoData &cfvo);
	bool readCfRule(QXmlStreamReader &reader, XlsxCfRuleData *cfRule, Styles *styles);
	bool readCfDataBar(QXmlStreamReader &reader, XlsxCfRuleData *cfRule);
	bool readCfColorScale(QXmlStreamReader &reader, XlsxCfRuleData *cfRule);

	QList<QSharedPointer<XlsxCfRuleData>> cfRules;
	QList<CellRange> ranges;
};

QT_END_NAMESPACE_XLSX

Q_DECLARE_METATYPE(QXlsx::XlsxCfVoData)
