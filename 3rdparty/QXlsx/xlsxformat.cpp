#include <QDataStream>
#include <QDebug>
#include <QtGlobal>

#include "xlsxcolor_p.h"
#include "xlsxformat.h"
#include "xlsxnumformatparser_p.h"

QT_BEGIN_NAMESPACE_XLSX

const QVector<FormatPrivate::Property> FormatPrivate::propertiesFont {
    FormatPrivate::Property::FontSize,
    FormatPrivate::Property::FontItalic,
    FormatPrivate::Property::FontStrikeOut,
    FormatPrivate::Property::FontColor,
    FormatPrivate::Property::FontBold,
    FormatPrivate::Property::FontScript,
    FormatPrivate::Property::FontUnderline,
    FormatPrivate::Property::FontOutline,
    FormatPrivate::Property::FontShadow,
    FormatPrivate::Property::FontName,
    FormatPrivate::Property::FontFamily,
    FormatPrivate::Property::FontCharset,
    FormatPrivate::Property::FontScheme,
    FormatPrivate::Property::FontCondense,
    FormatPrivate::Property::FontExtend,
};

const QVector<FormatPrivate::Property> FormatPrivate::propertiesBorder {
    FormatPrivate::Property::BorderLeftStyle,
    FormatPrivate::Property::BorderRightStyle,
    FormatPrivate::Property::BorderTopStyle,
    FormatPrivate::Property::BorderBottomStyle,
    FormatPrivate::Property::BorderDiagonalStyle,
    FormatPrivate::Property::BorderLeftColor,
    FormatPrivate::Property::BorderRightColor,
    FormatPrivate::Property::BorderTopColor,
    FormatPrivate::Property::BorderBottomColor,
    FormatPrivate::Property::BorderDiagonalColor,
    FormatPrivate::Property::BorderDiagonalType,
};

const QVector<FormatPrivate::Property> FormatPrivate::propertiesFill {
    FormatPrivate::Property::FillPattern,
    FormatPrivate::Property::FillBackgroundColor,
    FormatPrivate::Property::FillForegroundColor,
};

const QVector<FormatPrivate::Property> FormatPrivate::propertiesAlignment {
    FormatPrivate::Property::AlignmentHorizontal,
    FormatPrivate::Property::AlignmentVertical,
    FormatPrivate::Property::AlignmentWrap,
    FormatPrivate::Property::AlignmentRotation,
    FormatPrivate::Property::AlignmentIndent,
    FormatPrivate::Property::AlignmentShinkToFit,
};

FormatPrivate::FormatPrivate()
    : dirty(true)
    , font_dirty(true)
    , font_index_valid(false)
    , font_index(0)
    , fill_dirty(true)
    , fill_index_valid(false)
    , fill_index(0)
    , border_dirty(true)
    , border_index_valid(false)
    , border_index(0)
    , xf_index(-1)
    , xf_indexValid(false)
    , is_dxf_fomat(false)
    , dxf_index(-1)
    , dxf_indexValid(false)
    , theme(0)
{
}

FormatPrivate::FormatPrivate(const FormatPrivate &other)
    : QSharedData(other)
    , dirty(other.dirty)
    , formatKey(other.formatKey)
    , font_dirty(other.font_dirty)
    , font_index_valid(other.font_index_valid)
    , font_key(other.font_key)
    , font_index(other.font_index)
    , fill_dirty(other.fill_dirty)
    , fill_index_valid(other.fill_index_valid)
    , fill_key(other.fill_key)
    , fill_index(other.fill_index)
    , border_dirty(other.border_dirty)
    , border_index_valid(other.border_index_valid)
    , border_key(other.border_key)
    , border_index(other.border_index)
    , xf_index(other.xf_index)
    , xf_indexValid(other.xf_indexValid)
    , is_dxf_fomat(other.is_dxf_fomat)
    , dxf_index(other.dxf_index)
    , dxf_indexValid(other.dxf_indexValid)
    , theme(other.theme)
    , properties(other.properties)
{
}

FormatPrivate::~FormatPrivate()
{
}

QString FormatPrivate::toString() const
{
	QString out {"{"};

	const auto appendProperty = [&out, this](Property prop) {
		if (properties.contains(prop)) {
			out.push_back(QString {"{ %s, %s}, "}.arg(QVariant::fromValue(prop).toString()).arg(properties[prop].toString()));
		}
	};

	const auto appendProperties = [appendProperty](const QVector<Property> &props) {
		for (auto p : props) {
			appendProperty(p);
		}
	};

	appendProperty(Property::NumFmtId);
	appendProperty(Property::NumFmtFormatCode);

	appendProperties(propertiesFont);
	appendProperties(propertiesBorder);
	appendProperties(propertiesFill);
	appendProperties(propertiesAlignment);

	appendProperty(Property::ProtectionLocked);
	appendProperty(Property::ProtectionHidden);

	if (out.size() > 1) {
		out.remove(out.length() - 2, 1);  // removed last ', '
	}
	out.append("}");

	return out;
}
/*!
 * \class Format
 * \inmodule QtXlsx
 * \brief Providing the methods and properties that are available for formatting cells in Excel.
 */

/*!
 * \enum Format::FontScript
 *
 * The enum type defines the type of font script.
 *
 * \value Normal normal
 * \value Super super script
 * \value Sub sub script
 */

/*!
 * \enum Format::FontUnderline
 *
 * The enum type defines the type of font underline.
 *
 * \value FontUnderline::None
 * \value FontUnderline::Single
 * \value FontUnderline::Double
 * \value FontUnderline::SingleAccounting
 * \value FontUnderline::DoubleAccounting
 */

/*!
 * \enum Format::HorizontalAlignment
 *
 * The enum type defines the type of horizontal alignment.
 *
 * \value HorizontalAlignment::General
 * \value HorizontalAlignment::Left
 * \value HorizontalAlignment::Center
 * \value HorizontalAlignment::Right
 * \value HorizontalAlignment::Fill
 * \value HorizontalAlignment::Justify
 * \value HorizontalAlignment::Merge
 * \value HorizontalAlignment::Distributed
 */

/*!
 * \enum Format::VerticalAlignment
 *
 * The enum type defines the type of vertical alignment.
 *
 * \value VerticalAlignment::Top,
 * \value VerticalAlignment::Center,
 * \value VerticalAlignment::Bottom,
 * \value VerticalAlignment::Justify,
 * \value VerticalAlignment::Distributed
 */

/*!
 * \enum Format::BorderStyle
 *
 * The enum type defines the type of font underline.
 *
 * \value BorderStyle::None
 * \value BorderStyle::Thin
 * \value BorderStyle::Medium
 * \value BorderStyle::Dashed
 * \value BorderStyle::Dotted
 * \value BorderStyle::Thick
 * \value BorderStyle::Double
 * \value BorderStyle::Hair
 * \value BorderStyle::MediumDashed
 * \value BorderStyle::DashDot
 * \value BorderStyle::MediumDashDot
 * \value BorderStyle::DashDotDot
 * \value BorderStyle::MediumDashDotDot
 * \value BorderStyle::SlantDashDot
 */

/*!
 * \enum Format::DiagonalBorder
 *
 * The enum type defines the type of diagonal border.
 *
 * \value DiagonalBorder::None
 * \value DiagonalBorder::Down
 * \value DiagonalBorder::Up
 * \value DiagonalBorder::Both
 */

/*!
 * \enum Format::Fill
 *
 * The enum type defines the type of fill.
 *
 * \value Fill::None
 * \value Fill::Solid
 * \value Fill::MediumGray
 * \value Fill::DarkGray
 * \value Fill::LightGray
 * \value Fill::DarkHorizontal
 * \value Fill::DarkVertical
 * \value Fill::DarkDown
 * \value Fill::DarkUp
 * \value Fill::DarkGrid
 * \value Fill::DarkTrellis
 * \value Fill::LightHorizontal
 * \value Fill::LightVertical
 * \value Fill::LightDown
 * \value Fill::LightUp
 * \value Fill::LightTrellis
 * \value Fill::Gray125
 * \value Fill::Gray0625
 * \value Fill::LightGrid
 */

/*!
 *  Creates a new invalid format.
 */
Format::Format()
{
	// The d pointer is initialized with a null pointer
}

/*!
   Creates a new format with the same attributes as the \a other format.
 */
Format::Format(const Format &other)
    : d(other.d)
{
}

/*!
   Assigns the \a other format to this format, and returns a
   reference to this format.
 */
Format &Format::operator=(const Format &other)
{
	d = other.d;
	return *this;
}

/*!
 * Destroys this format.
 */
Format::~Format()
{
}

/*!
 * Returns the number format identifier.
 */
int Format::numberFormatIndex() const
{
	return intProperty(FormatPrivate::Property::NumFmtId, 0);
}

/*!
 * Set the number format identifier. The \a format
 * must be a valid built-in number format identifier
 * or the identifier of a custom number format.
 */
void Format::setNumberFormatIndex(int format)
{
	setProperty(FormatPrivate::Property::NumFmtId, QVariant::fromValue(format));
	clearProperty(FormatPrivate::Property::NumFmtFormatCode);
}

/*!
 * Returns the number format string.
 * \note for built-in number formats, this may
 * return an empty string.
 */
QString Format::numberFormat() const
{
	return stringProperty(FormatPrivate::Property::NumFmtFormatCode);
}

/*!
 * Set number \a format.
 * http://office.microsoft.com/en-001/excel-help/create-a-custom-number-format-HP010342372.aspx
 */
void Format::setNumberFormat(const QString &format)
{
	if (format.isEmpty())
		return;
	setProperty(FormatPrivate::Property::NumFmtFormatCode, format);
	clearProperty(FormatPrivate::Property::NumFmtId);  // numFmt id must be re-generated.
}

/*!
 * Returns whether the number format is probably a dateTime or not
 */
bool Format::isDateTimeFormat() const
{
	// NOTICE:

	if (hasProperty(FormatPrivate::Property::NumFmtFormatCode)) {
		// Custom numFmt, so
		// Gauss from the number string
		return NumFormatParser::isDateTime(numberFormat());
	} else if (hasProperty(FormatPrivate::Property::NumFmtId)) {
		// Non-custom numFmt
		int idx = numberFormatIndex();

		// Is built-in date time number id?
		if ((idx >= 14 && idx <= 22) || (idx >= 45 && idx <= 47))
			return true;

		if ((idx >= 27 && idx <= 36) || (idx >= 50 && idx <= 58))  // Used in CHS\CHT\JPN\KOR
			return true;
	}

	return false;
}

/*!
    \internal
    Set a custom num \a format with the given \a id.
 */
void Format::setNumberFormat(int id, const QString &format)
{
	setProperty(FormatPrivate::Property::NumFmtId, id);
	setProperty(FormatPrivate::Property::NumFmtFormatCode, format);
}

/*!
    \internal
    Called by styles to fix the numFmt
 */
void Format::fixNumberFormat(int id, const QString &format)
{
	setProperty(FormatPrivate::Property::NumFmtId, id, 0, false);
	setProperty(FormatPrivate::Property::NumFmtFormatCode, format, QString(), false);
}

/*!
    \internal
    Return true if the format has number format.
 */
bool Format::hasNumFmtData() const
{
	if (!d)
		return false;

	if (hasProperty(FormatPrivate::Property::NumFmtId) || hasProperty(FormatPrivate::Property::NumFmtFormatCode)) {
		return true;
	}
	return false;
}

/*!
 * Return the size of the font in points.
 */
int Format::fontSize() const
{
	return intProperty(FormatPrivate::Property::FontSize);
}

/*!
 * Set the \a size of the font in points.
 */
void Format::setFontSize(int size)
{
	setProperty(FormatPrivate::Property::FontSize, size, 0);
}

/*!
 * Return whether the font is italic.
 */
bool Format::fontItalic() const
{
	return boolProperty(FormatPrivate::Property::FontItalic);
}

/*!
 * Turn on/off the italic font based on \a italic.
 */
void Format::setFontItalic(bool italic)
{
	setProperty(FormatPrivate::Property::FontItalic, italic, false);
}

/*!
 * Return whether the font is strikeout.
 */
bool Format::fontStrikeOut() const
{
	return boolProperty(FormatPrivate::Property::FontStrikeOut);
}

/*!
 * Turn on/off the strikeOut font based on \a strikeOut.
 */
void Format::setFontStrikeOut(bool strikeOut)
{
	setProperty(FormatPrivate::Property::FontStrikeOut, strikeOut, false);
}

/*!
 * Return the color of the font.
 */
QColor Format::fontColor() const
{
	if (hasProperty(FormatPrivate::Property::FontColor))
		return colorProperty(FormatPrivate::Property::FontColor);
	return QColor();
}

/*!
 * Set the \a color of the font.
 */
void Format::setFontColor(const QColor &color)
{
	setProperty(FormatPrivate::Property::FontColor, XlsxColor(color), XlsxColor());
}

/*!
 * Return whether the font is bold.
 */
bool Format::fontBold() const
{
	return boolProperty(FormatPrivate::Property::FontBold);
}

/*!
 * Turn on/off the bold font based on the given \a bold.
 */
void Format::setFontBold(bool bold)
{
	setProperty(FormatPrivate::Property::FontBold, bold, false);
}

/*!
 * Return the script style of the font.
 */
Format::FontScript Format::fontScript() const
{
	return static_cast<Format::FontScript>(intProperty(FormatPrivate::Property::FontScript));
}

/*!
 * Set the script style of the font to \a script.
 */
void Format::setFontScript(FontScript script)
{
	setProperty(FormatPrivate::Property::FontScript, QVariant::fromValue(script), QVariant::fromValue(FontScript::Normal));
}

/*!
 * Return the underline style of the font.
 */
Format::FontUnderline Format::fontUnderline() const
{
	return static_cast<Format::FontUnderline>(intProperty(FormatPrivate::Property::FontUnderline));
}

/*!
 * Set the underline style of the font to \a underline.
 */
void Format::setFontUnderline(FontUnderline underline)
{
	setProperty(FormatPrivate::Property::FontUnderline, QVariant::fromValue(underline), QVariant::fromValue(FontUnderline::None));
}

/*!
 * Return whether the font is outline.
 */
bool Format::fontOutline() const
{
	return boolProperty(FormatPrivate::Property::FontOutline);
}

/*!
 * Turn on/off the outline font based on \a outline.
 */
void Format::setFontOutline(bool outline)
{
	setProperty(FormatPrivate::Property::FontOutline, outline, false);
}

/*!
 * Return the name of the font.
 */
QString Format::fontName() const
{
	return stringProperty(FormatPrivate::Property::FontName, QStringLiteral("Calibri"));
}

/*!
 * Set the name of the font to \a name.
 */
void Format::setFontName(const QString &name)
{
	setProperty(FormatPrivate::Property::FontName, name, QStringLiteral("Calibri"));
}

/*!
 * Returns a QFont object based on font data contained in the format.
 */
QFont Format::font() const
{
	QFont font;
	font.setFamily(fontName());
	if (fontSize() > 0)
		font.setPointSize(fontSize());
	font.setBold(fontBold());
	font.setItalic(fontItalic());
	font.setUnderline(fontUnderline() != FontUnderline::None);
	font.setStrikeOut(fontStrikeOut());
	return font;
}

/*!
 * Set the format properties from the given \a font.
 */
void Format::setFont(const QFont &font)
{
	setFontName(font.family());
	if (font.pointSize() > 0)
		setFontSize(font.pointSize());
	setFontBold(font.bold());
	setFontItalic(font.italic());
	setFontUnderline(font.underline() ? FontUnderline::Single : FontUnderline::None);
	setFontStrikeOut(font.strikeOut());
}

/*!
 * \internal
 * When the format has font data, when need to assign a valid index for it.
 * The index value is depend on the order <fonts > in styles.xml
 */
bool Format::fontIndexValid() const
{
	if (!hasFontData())
		return false;
	return d->font_index_valid;
}

/*!
 * \internal
 */
int Format::fontIndex() const
{
	if (fontIndexValid())
		return d->font_index;

	return 0;
}

/*!
 * \internal
 */
void Format::setFontIndex(int index)
{
	d->font_index = index;
	d->font_index_valid = true;
}

/*!
 * \internal
 */
QByteArray Format::fontKey() const
{
	if (isEmpty())
		return QByteArray();

	if (d->font_dirty) {
		QByteArray key;
		QDataStream stream(&key, QIODevice::WriteOnly);
		for (int i = 0; i < d->propertiesFont.size(); ++i) {
			auto it = d->properties.constFind(d->propertiesFont.at(i));
			if (it != d->properties.constEnd())
				stream << i << it.value();
		};

		const_cast<Format *>(this)->d->font_key = key;
		const_cast<Format *>(this)->d->font_dirty = false;
	}

	return d->font_key;
}

/*!
    \internal
    Return true if the format has font format, otherwise return false.
 */
bool Format::hasFontData() const
{
	if (!d)
		return false;

	for (int i = 0; i < d->propertiesFont.size(); ++i) {
		if (hasProperty(d->propertiesFont.at(i)))
			return true;
	}
	return false;
}

/*!
 * Return the horizontal alignment.
 */
Format::HorizontalAlignment Format::horizontalAlignment() const
{
	return static_cast<Format::HorizontalAlignment>(
	    intProperty(FormatPrivate::Property::AlignmentHorizontal, QVariant::fromValue(HorizontalAlignment::General).toInt()));
}

/*!
 * Set the horizontal alignment with the given \a align.
 */
void Format::setHorizontalAlignment(HorizontalAlignment align)
{
	if (hasProperty(FormatPrivate::Property::AlignmentIndent)
	    && (align != HorizontalAlignment::General && align != HorizontalAlignment::Left && align != HorizontalAlignment::Right
	        && align != HorizontalAlignment::Distributed)) {
		clearProperty(FormatPrivate::Property::AlignmentIndent);
	}

	if (hasProperty(FormatPrivate::Property::AlignmentShinkToFit)
	    && (align == HorizontalAlignment::Fill || align == HorizontalAlignment::Justify || align == HorizontalAlignment::Distributed)) {
		clearProperty(FormatPrivate::Property::AlignmentShinkToFit);
	}

	setProperty(FormatPrivate::Property::AlignmentHorizontal, QVariant::fromValue(align), QVariant::fromValue(HorizontalAlignment::General));
}

/*!
 * Return the vertical alignment.
 */
Format::VerticalAlignment Format::verticalAlignment() const
{
	return static_cast<Format::VerticalAlignment>(
	    intProperty(FormatPrivate::Property::AlignmentVertical, QVariant::fromValue(VerticalAlignment::Bottom).toInt()));
}

/*!
 * Set the vertical alignment with the given \a align.
 */
void Format::setVerticalAlignment(VerticalAlignment align)
{
	setProperty(FormatPrivate::Property::AlignmentVertical, QVariant::fromValue(align), QVariant::fromValue(VerticalAlignment::Bottom));
}

/*!
 * Return whether the cell text is wrapped.
 */
bool Format::textWrap() const
{
	return boolProperty(FormatPrivate::Property::AlignmentWrap);
}

/*!
 * Enable the text wrap if \a wrap is true.
 */
void Format::setTextWrap(bool wrap)
{
	if (wrap && hasProperty(FormatPrivate::Property::AlignmentShinkToFit))
		clearProperty(FormatPrivate::Property::AlignmentShinkToFit);

	setProperty(FormatPrivate::Property::AlignmentWrap, wrap, false);
}

/*!
 * Return the text rotation.
 */
int Format::rotation() const
{
	return intProperty(FormatPrivate::Property::AlignmentRotation);
}

/*!
 * Set the text roation with the given \a rotation. Must be in the range [0, 180] or 255.
 */
void Format::setRotation(int rotation)
{
	setProperty(FormatPrivate::Property::AlignmentRotation, rotation, 0);
}

/*!
 * Return the text indentation level.
 */
int Format::indent() const
{
	return intProperty(FormatPrivate::Property::AlignmentIndent);
}

/*!
 * Set the text indentation level with the given \a indent. Must be less than or equal to 15.
 */
void Format::setIndent(int indent)
{
	if (indent && hasProperty(FormatPrivate::Property::AlignmentHorizontal)) {
		HorizontalAlignment hl = horizontalAlignment();

		if (hl != HorizontalAlignment::General && hl != HorizontalAlignment::Left && hl != HorizontalAlignment::Right && hl != HorizontalAlignment::Justify) {
			setHorizontalAlignment(HorizontalAlignment::Left);
		}
	}

	setProperty(FormatPrivate::Property::AlignmentIndent, indent, 0);
}

/*!
 * Return whether the cell is shrink to fit.
 */
bool Format::shrinkToFit() const
{
	return boolProperty(FormatPrivate::Property::AlignmentShinkToFit);
}

/*!
 * Turn on/off shrink to fit base on \a shink.
 */
void Format::setShrinkToFit(bool shink)
{
	if (shink && hasProperty(FormatPrivate::Property::AlignmentWrap))
		clearProperty(FormatPrivate::Property::AlignmentWrap);

	if (shink && hasProperty(FormatPrivate::Property::AlignmentHorizontal)) {
		HorizontalAlignment hl = horizontalAlignment();
		if (hl == HorizontalAlignment::Fill || hl == HorizontalAlignment::Justify || hl == HorizontalAlignment::Distributed)
			setHorizontalAlignment(HorizontalAlignment::Left);
	}

	setProperty(FormatPrivate::Property::AlignmentShinkToFit, shink, false);
}

/*!
 * \internal
 */
bool Format::hasAlignmentData() const
{
	if (!d)
		return false;

	for (int i = 0; i < d->propertiesAlignment.size(); ++i) {
		if (hasProperty(d->propertiesAlignment.at(0)))
			return true;
	}
	return false;
}

/*!
 * Set the border style with the given \a style.
 */
void Format::setBorderStyle(BorderStyle style)
{
	setLeftBorderStyle(style);
	setRightBorderStyle(style);
	setBottomBorderStyle(style);
	setTopBorderStyle(style);
}

/*!
 * Sets the border color with the given \a color.
 */
void Format::setBorderColor(const QColor &color)
{
	setLeftBorderColor(color);
	setRightBorderColor(color);
	setTopBorderColor(color);
	setBottomBorderColor(color);
}

/*!
 * Returns the left border style
 */
Format::BorderStyle Format::leftBorderStyle() const
{
	return static_cast<BorderStyle>(intProperty(FormatPrivate::Property::BorderLeftStyle));
}

/*!
 * Sets the left border style to \a style
 */
void Format::setLeftBorderStyle(BorderStyle style)
{
	setProperty(FormatPrivate::Property::BorderLeftStyle, QVariant::fromValue(style), QVariant::fromValue(BorderStyle::None));
}

/*!
 * Returns the left border color
 */
QColor Format::leftBorderColor() const
{
	return colorProperty(FormatPrivate::Property::BorderLeftColor);
}

/*!
    Sets the left border color to the given \a color
*/
void Format::setLeftBorderColor(const QColor &color)
{
	setProperty(FormatPrivate::Property::BorderLeftColor, XlsxColor(color), XlsxColor());
}

/*!
    Returns the right border style.
*/
Format::BorderStyle Format::rightBorderStyle() const
{
	return static_cast<BorderStyle>(intProperty(FormatPrivate::Property::BorderRightStyle));
}

/*!
    Sets the right border style to the given \a style.
*/
void Format::setRightBorderStyle(BorderStyle style)
{
	setProperty(FormatPrivate::Property::BorderRightStyle, QVariant::fromValue(style), QVariant::fromValue(BorderStyle::None));
}

/*!
    Returns the right border color.
*/
QColor Format::rightBorderColor() const
{
	return colorProperty(FormatPrivate::Property::BorderRightColor);
}

/*!
    Sets the right border color to the given \a color
*/
void Format::setRightBorderColor(const QColor &color)
{
	setProperty(FormatPrivate::Property::BorderRightColor, XlsxColor(color), XlsxColor());
}

/*!
    Returns the top border style.
*/
Format::BorderStyle Format::topBorderStyle() const
{
	return static_cast<BorderStyle>(intProperty(FormatPrivate::Property::BorderTopStyle));
}

/*!
    Sets the top border style to the given \a style.
*/
void Format::setTopBorderStyle(BorderStyle style)
{
	setProperty(FormatPrivate::Property::BorderTopStyle, QVariant::fromValue(style), QVariant::fromValue(BorderStyle::None));
}

/*!
    Returns the top border color.
*/
QColor Format::topBorderColor() const
{
	return colorProperty(FormatPrivate::Property::BorderTopColor);
}

/*!
    Sets the top border color to the given \a color.
*/
void Format::setTopBorderColor(const QColor &color)
{
	setProperty(FormatPrivate::Property::BorderTopColor, XlsxColor(color), XlsxColor());
}

/*!
    Returns the bottom border style.
*/
Format::BorderStyle Format::bottomBorderStyle() const
{
	return static_cast<BorderStyle>(intProperty(FormatPrivate::Property::BorderBottomStyle));
}

/*!
    Sets the bottom border style to the given \a style.
*/
void Format::setBottomBorderStyle(BorderStyle style)
{
	setProperty(FormatPrivate::Property::BorderBottomStyle, QVariant::fromValue(style), QVariant::fromValue(BorderStyle::None));
}

/*!
    Returns the bottom border color.
*/
QColor Format::bottomBorderColor() const
{
	return colorProperty(FormatPrivate::Property::BorderBottomColor);
}

/*!
    Sets the bottom border color to the given \a color.
*/
void Format::setBottomBorderColor(const QColor &color)
{
	setProperty(FormatPrivate::Property::BorderBottomColor, XlsxColor(color), XlsxColor());
}

/*!
    Return the diagonla border style.
*/
Format::BorderStyle Format::diagonalBorderStyle() const
{
	return static_cast<BorderStyle>(intProperty(FormatPrivate::Property::BorderDiagonalStyle));
}

/*!
    Sets the diagonal border style to the given \a style.
*/
void Format::setDiagonalBorderStyle(BorderStyle style)
{
	setProperty(FormatPrivate::Property::BorderDiagonalStyle, QVariant::fromValue(style), QVariant::fromValue(BorderStyle::None));
}

/*!
    Returns the diagonal border type.
*/
Format::DiagonalBorder Format::diagonalBorderType() const
{
	return static_cast<DiagonalBorder>(intProperty(FormatPrivate::Property::BorderDiagonalType));
}

/*!
    Sets the diagonal border type to the given \a style
*/
void Format::setDiagonalBorderType(DiagonalBorder style)
{
	setProperty(FormatPrivate::Property::BorderDiagonalType, QVariant::fromValue(style), QVariant::fromValue(DiagonalBorder::None));
}

/*!
    Returns the diagonal border color.
*/
QColor Format::diagonalBorderColor() const
{
	return colorProperty(FormatPrivate::Property::BorderDiagonalColor);
}

/*!
    Sets the diagonal border color to the given \a color
*/
void Format::setDiagonalBorderColor(const QColor &color)
{
	setProperty(FormatPrivate::Property::BorderDiagonalColor, XlsxColor(color), XlsxColor());
}

/*!
    \internal
    Returns whether this format has been set valid border index.
*/
bool Format::borderIndexValid() const
{
	if (!hasBorderData())
		return false;
	return d->border_index_valid;
}

/*!
    \internal
    Returns the border index.
*/
int Format::borderIndex() const
{
	if (borderIndexValid())
		return d->border_index;
	return 0;
}

/*!
 * \internal
 */
void Format::setBorderIndex(int index)
{
	d->border_index = index;
	d->border_index_valid = true;
}

/*! \internal
 */
QByteArray Format::borderKey() const
{
	if (isEmpty())
		return QByteArray();

	if (d->border_dirty) {
		QByteArray key;
		QDataStream stream(&key, QIODevice::WriteOnly);
		for (int i = 0; i < d->propertiesBorder.size(); ++i) {
			auto it = d->properties.constFind(d->propertiesBorder.at(i));
			if (it != d->properties.constEnd())
				stream << i << it.value();
		};

		const_cast<Format *>(this)->d->border_key = key;
		const_cast<Format *>(this)->d->border_dirty = false;
	}

	return d->border_key;
}

/*!
    \internal
    Return true if the format has border format, otherwise return false.
 */
bool Format::hasBorderData() const
{
	if (!d)
		return false;

	for (auto prop : d->propertiesBorder) {
		if (hasProperty(prop))
			return true;
	}
	return false;
}

/*!
    Return the fill pattern.
*/
Format::Fill Format::fillPattern() const
{
	return static_cast<Fill>(intProperty(FormatPrivate::Property::FillPattern, QVariant::fromValue(Fill::None).toInt()));
}

/*!
    Sets the fill pattern to the given \a pattern.
*/
void Format::setFillPattern(Fill pattern)
{
	setProperty(FormatPrivate::Property::FillPattern, QVariant::fromValue(pattern), QVariant::fromValue(Fill::None));
}

/*!
    Returns the foreground color of the pattern.
*/
QColor Format::patternForegroundColor() const
{
	return colorProperty(FormatPrivate::Property::FillForegroundColor);
}

/*!
    Sets the foreground color of the pattern with the given \a color.
*/
void Format::setPatternForegroundColor(const QColor &color)
{
	if (color.isValid() && !hasProperty(FormatPrivate::Property::FillPattern))
		setFillPattern(Fill::Solid);
	setProperty(FormatPrivate::Property::FillForegroundColor, XlsxColor(color), XlsxColor());
}

/*!
    Returns the background color of the pattern.
*/
QColor Format::patternBackgroundColor() const
{
	return colorProperty(FormatPrivate::Property::FillBackgroundColor);
}

/*!
    Sets the background color of the pattern with the given \a color.
*/
void Format::setPatternBackgroundColor(const QColor &color)
{
	if (color.isValid() && !hasProperty(FormatPrivate::Property::FillPattern))
		setFillPattern(Fill::Solid);
	setProperty(FormatPrivate::Property::FillBackgroundColor, XlsxColor(color), XlsxColor());
}

/*!
 * \internal
 */
bool Format::fillIndexValid() const
{
	if (!hasFillData())
		return false;
	return d->fill_index_valid;
}

/*!
 * \internal
 */
int Format::fillIndex() const
{
	if (fillIndexValid())
		return d->fill_index;
	return 0;
}

/*!
 * \internal
 */
void Format::setFillIndex(int index)
{
	d->fill_index = index;
	d->fill_index_valid = true;
}

/*!
 * \internal
 */
QByteArray Format::fillKey() const
{
	if (isEmpty())
		return QByteArray();

	if (d->fill_dirty) {
		QByteArray key;
		QDataStream stream(&key, QIODevice::WriteOnly);
		for (int i = 0; i < d->propertiesFill.size(); ++i) {
			auto it = d->properties.constFind(d->propertiesFill.at(i));
			if (it != d->properties.constEnd())
				stream << i << it.value().toString();
		};

		const_cast<Format *>(this)->d->fill_key = key;
		const_cast<Format *>(this)->d->fill_dirty = false;
	}

	return d->fill_key;
}

/*!
    \internal
    Return true if the format has fill format, otherwise return false.
 */
bool Format::hasFillData() const
{
	if (!d)
		return false;

	for (auto prop : d->propertiesFill) {
		if (hasProperty(prop)) {
			return true;
		}
	}
	return false;
}

/*!
    Returns whether the hidden protection property is set to true.
*/
bool Format::hidden() const
{
	return boolProperty(FormatPrivate::Property::ProtectionHidden);
}

/*!
    Sets the hidden protection property with the given \a hidden.
*/
void Format::setHidden(bool hidden)
{
	setProperty(FormatPrivate::Property::ProtectionHidden, hidden);
}

/*!
    Returns whether the locked protection property is set to true.
*/
bool Format::locked() const
{
	return boolProperty(FormatPrivate::Property::ProtectionLocked);
}

/*!
    Sets the locked protection property with the given \a locked.
*/
void Format::setLocked(bool locked)
{
	setProperty(FormatPrivate::Property::ProtectionLocked, locked);
}

/*!
    \internal
    Return true if the format has protection data, otherwise return false.
 */
bool Format::hasProtectionData() const
{
	if (!d)
		return false;

	if (hasProperty(FormatPrivate::Property::ProtectionHidden) || hasProperty(FormatPrivate::Property::ProtectionLocked)) {
		return true;
	}
	return false;
}

/*!
    Merges the current format with the properties described by format \a modifier.
 */
void Format::mergeFormat(const Format &modifier)
{
	if (!modifier.isValid())
		return;

	if (!isValid()) {
		d = modifier.d;
		return;
	}

	QMapIterator<FormatPrivate::Property, QVariant> it(modifier.d->properties);
	while (it.hasNext()) {
		it.next();
		setProperty(it.key(), it.value());
	}
}

/*!
    Returns true if the format is valid; otherwise returns false.
 */
bool Format::isValid() const
{
	if (d)
		return true;
	return false;
}

/*!
    Returns true if the format is empty; otherwise returns false.
 */
bool Format::isEmpty() const
{
	if (!d)
		return true;
	return d->properties.isEmpty();
}

/*!
 * \internal
 */
QByteArray Format::formatKey() const
{
	if (isEmpty())
		return QByteArray();

	if (d->dirty) {
		QByteArray key;
		QDataStream stream(&key, QIODevice::WriteOnly);

		QMapIterator<FormatPrivate::Property, QVariant> i(d->properties);
		while (i.hasNext()) {
			i.next();
			stream << i.key() << i.value();
		}

		d->formatKey = key;
		d->dirty = false;
	}

	return d->formatKey;
}

/*!
 * \internal
 *  Called by QXlsx::Styles or some unittests.
 */
void Format::setXfIndex(int index)
{
	if (!d)
		d = new FormatPrivate;
	d->xf_index = index;
	d->xf_indexValid = true;
}

/*!
 * \internal
 */
int Format::xfIndex() const
{
	if (!d)
		return -1;
	return d->xf_index;
}

/*!
 * \internal
 */
bool Format::xfIndexValid() const
{
	if (!d)
		return false;
	return d->xf_indexValid;
}

/*!
 *  Called by QXlsx::Styles or some unittests.
 */
void Format::setDxfIndex(int index)
{
	if (!d)
		d = new FormatPrivate;
	d->dxf_index = index;
	d->dxf_indexValid = true;
}

int Format::dxfIndex() const
{
	if (!d)
		return -1;
	return d->dxf_index;
}

bool Format::dxfIndexValid() const
{
	if (!d)
		return false;
	return d->dxf_indexValid;
}

/*!
    Returns ture if the \a format is equal to this format.
*/
bool Format::operator==(const Format &format) const
{
	return this->formatKey() == format.formatKey();
}

/*!
    Returns ture if the \a format is not equal to this format.
*/
bool Format::operator!=(const Format &format) const
{
	return this->formatKey() != format.formatKey();
}

int Format::theme() const
{
	return d->theme;
}

QVariant Format::property(FormatPrivate::Property propertyId, const QVariant &defaultValue) const
{
	if (d) {
		auto it = d->properties.constFind(propertyId);
		if (it != d->properties.constEnd())
			return it.value();
	}
	return defaultValue;
}

void Format::setProperty(FormatPrivate::Property propertyId, const QVariant &value, const QVariant &clearValue, bool detach)
{
	if (!d)
		d = new FormatPrivate;

	if (value != clearValue) {
		auto it = d->properties.constFind(propertyId);
		if (it != d->properties.constEnd() && it.value() == value)
			return;

		if (detach)
			d.detach();

		d->properties[propertyId] = value;
	} else {
		if (!d->properties.contains(propertyId))
			return;

		if (detach)
			d.detach();

		d->properties.remove(propertyId);
	}

	d->dirty = true;
	d->xf_indexValid = false;
	d->dxf_indexValid = false;

	if (d->propertiesFont.contains(propertyId)) {
		d->font_dirty = true;
		d->font_index_valid = false;
	} else if (d->propertiesBorder.contains(propertyId)) {
		d->border_dirty = true;
		d->border_index_valid = false;
	} else if (d->propertiesFill.contains(propertyId)) {
		d->fill_dirty = true;
		d->fill_index_valid = false;
	}
}

/*!
 * \internal
 */
void Format::clearProperty(FormatPrivate::Property propertyId)
{
	setProperty(propertyId, {});
}

/*!
 * \internal
 */
bool Format::hasProperty(FormatPrivate::Property propertyId) const
{
	if (!d)
		return false;
	return d->properties.contains(propertyId);
}

bool Format::boolProperty(FormatPrivate::Property propertyId, bool defaultValue) const
{
	if (!hasProperty(propertyId))
		return defaultValue;

	const QVariant prop = d->properties[propertyId];
	if (prop.userType() != QMetaType::Bool)
		return defaultValue;
	return prop.toBool();
}

int Format::intProperty(FormatPrivate::Property propertyId, const QVariant &defaultValue) const
{
	if (!hasProperty(propertyId))
		return defaultValue.toInt();

	const QVariant prop = d->properties[propertyId];
	if (prop.userType() != QMetaType::Int)
		return defaultValue.toInt();
	return prop.toInt();
}

double Format::doubleProperty(FormatPrivate::Property propertyId, double defaultValue) const
{
	if (!hasProperty(propertyId))
		return defaultValue;

	const QVariant prop = d->properties[propertyId];
	if (prop.userType() != QMetaType::Double && prop.userType() != QMetaType::Float)
		return defaultValue;
	return prop.toDouble();
}

QString Format::stringProperty(FormatPrivate::Property propertyId, const QString &defaultValue) const
{
	if (!hasProperty(propertyId))
		return defaultValue;

	const QVariant prop = d->properties[propertyId];
	if (prop.userType() != QMetaType::QString)
		return defaultValue;
	return prop.toString();
}

QColor Format::colorProperty(FormatPrivate::Property propertyId, const QColor &defaultValue) const
{
	if (!hasProperty(propertyId))
		return defaultValue;

	const QVariant prop = d->properties[propertyId];
	if (prop.userType() != qMetaTypeId<XlsxColor>())
		return defaultValue;
	return qvariant_cast<XlsxColor>(prop).rgbColor();
}

#ifndef QT_NO_DEBUG_STREAM
QDebug operator<<(QDebug dbg, const Format &f)
{
	dbg.nospace() << "QXlsx::Format(" << f.d->toString() << ")";
	return dbg.space();
}
#endif

QT_END_NAMESPACE_XLSX
