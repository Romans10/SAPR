#pragma once

#include <QByteArray>
#include <QColor>
#include <QExplicitlySharedDataPointer>
#include <QFont>
#include <QList>
#include <QMetaEnum>
#include <QVariant>

#include "xlsxformat_p.h"
#include "xlsxglobal.h"

class FormatTest;

QT_BEGIN_NAMESPACE_XLSX

class Styles;
class Worksheet;
class WorksheetPrivate;
class RichStringPrivate;
class SharedStrings;

class Format
{
	Q_GADGET
public:
	enum class FontScript
	{
		Normal,
		Super,
		Sub
	};
	Q_ENUM(FontScript)

	enum class FontUnderline
	{
		None,
		Single,
		Double,
		SingleAccounting,
		DoubleAccounting
	};
	Q_ENUM(FontUnderline)

	enum class HorizontalAlignment
	{
		General,
		Left,
		Center,
		Right,
		Fill,
		Justify,
		Merge,
		Distributed
	};
	Q_ENUM(HorizontalAlignment)

	enum class VerticalAlignment
	{
		Top,
		Center,
		Bottom,
		Justify,
		Distributed
	};
	Q_ENUM(VerticalAlignment)

	enum class BorderStyle
	{
		None,
		Thin,
		Medium,
		Dashed,
		Dotted,
		Thick,
		Double,
		Hair,
		MediumDashed,
		DashDot,
		MediumDashDot,
		DashDotDot,
		MediumDashDotDot,
		SlantDashDot
	};
	Q_ENUM(BorderStyle)

	enum class DiagonalBorder
	{
		None,
		Down,
		Up,
		Both
	};
	Q_ENUM(DiagonalBorder)

	enum class Fill
	{
		None,
		Solid,
		MediumGray,
		DarkGray,
		LightGray,
		DarkHorizontal,
		DarkVertical,
		DarkDown,
		DarkUp,
		DarkGrid,
		DarkTrellis,
		LightHorizontal,
		LightVertical,
		LightDown,
		LightUp,
		LightTrellis,
		Gray125,
		Gray0625,
		LightGrid
	};
	Q_ENUM(Fill)

	Format();
	Format(const Format &other);
	Format &operator=(const Format &rhs);
	~Format();

	int numberFormatIndex() const;
	void setNumberFormatIndex(int format);
	QString numberFormat() const;
	void setNumberFormat(const QString &format);
	void setNumberFormat(int id, const QString &format);
	bool isDateTimeFormat() const;

	int fontSize() const;
	void setFontSize(int size);
	bool fontItalic() const;
	void setFontItalic(bool italic);
	bool fontStrikeOut() const;
	void setFontStrikeOut(bool);
	QColor fontColor() const;
	void setFontColor(const QColor &);
	bool fontBold() const;
	void setFontBold(bool bold);
	FontScript fontScript() const;
	void setFontScript(FontScript);
	FontUnderline fontUnderline() const;
	void setFontUnderline(FontUnderline);
	bool fontOutline() const;
	void setFontOutline(bool outline);
	QString fontName() const;
	void setFontName(const QString &);
	QFont font() const;
	void setFont(const QFont &font);

	HorizontalAlignment horizontalAlignment() const;
	void setHorizontalAlignment(HorizontalAlignment align);
	VerticalAlignment verticalAlignment() const;
	void setVerticalAlignment(VerticalAlignment align);
	bool textWrap() const;
	void setTextWrap(bool textWrap);
	int rotation() const;
	void setRotation(int rotation);
	int indent() const;
	void setIndent(int indent);
	bool shrinkToFit() const;
	void setShrinkToFit(bool shink);

	void setBorderStyle(BorderStyle style);
	void setBorderColor(const QColor &color);
	BorderStyle leftBorderStyle() const;
	void setLeftBorderStyle(BorderStyle style);
	QColor leftBorderColor() const;
	void setLeftBorderColor(const QColor &color);
	BorderStyle rightBorderStyle() const;
	void setRightBorderStyle(BorderStyle style);
	QColor rightBorderColor() const;
	void setRightBorderColor(const QColor &color);
	BorderStyle topBorderStyle() const;
	void setTopBorderStyle(BorderStyle style);
	QColor topBorderColor() const;
	void setTopBorderColor(const QColor &color);
	BorderStyle bottomBorderStyle() const;
	void setBottomBorderStyle(BorderStyle style);
	QColor bottomBorderColor() const;
	void setBottomBorderColor(const QColor &color);
	BorderStyle diagonalBorderStyle() const;
	void setDiagonalBorderStyle(BorderStyle style);
	DiagonalBorder diagonalBorderType() const;
	void setDiagonalBorderType(DiagonalBorder style);
	QColor diagonalBorderColor() const;
	void setDiagonalBorderColor(const QColor &color);

	Fill fillPattern() const;
	void setFillPattern(Fill pattern);
	QColor patternForegroundColor() const;
	void setPatternForegroundColor(const QColor &color);
	QColor patternBackgroundColor() const;
	void setPatternBackgroundColor(const QColor &color);

	bool locked() const;
	void setLocked(bool locked);
	bool hidden() const;
	void setHidden(bool hidden);

	void mergeFormat(const Format &modifier);
	bool isValid() const;
	bool isEmpty() const;

	bool operator==(const Format &format) const;
	bool operator!=(const Format &format) const;

	QVariant property(FormatPrivate::Property propertyId, const QVariant &defaultValue = {}) const;
	void setProperty(FormatPrivate::Property propertyId, const QVariant &value, const QVariant &clearValue = {}, bool detach = true);
	void clearProperty(FormatPrivate::Property propertyId);
	bool hasProperty(FormatPrivate::Property propertyId) const;

	bool boolProperty(FormatPrivate::Property propertyId, bool defaultValue = false) const;
	int intProperty(FormatPrivate::Property propertyId, const QVariant &defaultValue = 0) const;
	double doubleProperty(FormatPrivate::Property propertyId, double defaultValue = 0.0) const;
	QString stringProperty(FormatPrivate::Property propertyId, const QString &defaultValue = QString()) const;
	QColor colorProperty(FormatPrivate::Property propertyId, const QColor &defaultValue = QColor()) const;

	bool hasNumFmtData() const;
	bool hasFontData() const;
	bool hasFillData() const;
	bool hasBorderData() const;
	bool hasAlignmentData() const;
	bool hasProtectionData() const;

	bool fontIndexValid() const;
	int fontIndex() const;
	QByteArray fontKey() const;
	bool borderIndexValid() const;
	QByteArray borderKey() const;
	int borderIndex() const;
	bool fillIndexValid() const;
	QByteArray fillKey() const;
	int fillIndex() const;

	QByteArray formatKey() const;
	bool xfIndexValid() const;
	int xfIndex() const;
	bool dxfIndexValid() const;
	int dxfIndex() const;

	void fixNumberFormat(int id, const QString &format);
	void setFontIndex(int index);
	void setBorderIndex(int index);
	void setFillIndex(int index);
	void setXfIndex(int index);
	void setDxfIndex(int index);

private:
	friend class Styles;
	friend class ::FormatTest;
	friend QDebug operator<<(QDebug, const Format &f);

	int theme() const;

	QExplicitlySharedDataPointer<FormatPrivate> d;
};

#ifndef QT_NO_DEBUG_STREAM
QDebug operator<<(QDebug dbg, const Format &f);
#endif

QT_END_NAMESPACE_XLSX
