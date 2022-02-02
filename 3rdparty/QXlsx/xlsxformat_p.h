#pragma once

#include <QMap>
#include <QSet>
#include <QSharedData>

#include "xlsxglobal.h"

QT_BEGIN_NAMESPACE_XLSX

class FormatPrivate : public QSharedData
{
	Q_GADGET
public:
	enum class Type
	{
		Invalid = 0,
		NumFmt = 0x01,
		Font = 0x02,
		Alignment = 0x04,
		Border = 0x08,
		Fill = 0x10,
		Protection = 0x20
	};

	enum class Property
	{
		// numFmt
		NumFmtId,
		NumFmtFormatCode,

		// font
		FontSize,
		FontItalic,
		FontStrikeOut,
		FontColor,
		FontBold,
		FontScript,
		FontUnderline,
		FontOutline,
		FontShadow,
		FontName,
		FontFamily,
		FontCharset,
		FontScheme,
		FontCondense,
		FontExtend,

		// border
		BorderLeftStyle,
		BorderRightStyle,
		BorderTopStyle,
		BorderBottomStyle,
		BorderDiagonalStyle,
		BorderLeftColor,
		BorderRightColor,
		BorderTopColor,
		BorderBottomColor,
		BorderDiagonalColor,
		BorderDiagonalType,

		// fill
		FillPattern,
		FillBackgroundColor,
		FillForegroundColor,

		// alignment
		AlignmentHorizontal,
		AlignmentVertical,
		AlignmentWrap,
		AlignmentRotation,
		AlignmentIndent,
		AlignmentShinkToFit,

		// protection
		ProtectionLocked,
		ProtectionHidden,
	};
	Q_ENUM(Property)

	FormatPrivate();
	FormatPrivate(const FormatPrivate &other);
	~FormatPrivate();

	QString toString() const;

	bool dirty;  // The key re-generation is need.
	QByteArray formatKey;

	bool font_dirty;
	bool font_index_valid;
	QByteArray font_key;
	int font_index;

	bool fill_dirty;
	bool fill_index_valid;
	QByteArray fill_key;
	int fill_index;

	bool border_dirty;
	bool border_index_valid;
	QByteArray border_key;
	int border_index;

	int xf_index;
	bool xf_indexValid;

	bool is_dxf_fomat;
	int dxf_index;
	bool dxf_indexValid;

	int theme;

	QMap<Property, QVariant> properties;
	static const QVector<Property> propertiesFont;
	static const QVector<Property> propertiesBorder;
	static const QVector<Property> propertiesFill;
	static const QVector<Property> propertiesAlignment;
};

QT_END_NAMESPACE_XLSX
