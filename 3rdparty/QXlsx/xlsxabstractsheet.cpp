#include <QtGlobal>

#include "xlsxabstractsheet.h"
#include "xlsxabstractsheet_p.h"
#include "xlsxworkbook.h"

QT_BEGIN_NAMESPACE_XLSX

AbstractSheetPrivate::AbstractSheetPrivate(AbstractSheet *p, AbstractSheet::CreateFlag flag)
    : AbstractXmlFilePrivate(p, flag)
{
	type = AbstractSheet::Type::Work;
	sheetState = AbstractSheet::State::Visible;
}

AbstractSheetPrivate::~AbstractSheetPrivate()
{
}

/*!
  \class AbstractSheet
  \inmodule QtXlsx
  \brief Base class for worksheet, chartsheet, etc.
*/

/*!
  \enum AbstractSheet::Type

  \value Work
  \value Chart
  \omitvalue Dialog
  \omitvalue Macro
*/

/*!
  \enum AbstractSheet::State

  \value Visible
  \value Hidden
  \value VeryHidden User cann't make a veryHidden sheet visible in normal way.
*/

/*!
  \fn AbstractSheet::copy(const QString &distName, int distId) const

  Copies the current sheet to a sheet called \a distName with \a distId.
  Returns the new sheet.
 */

AbstractSheet::AbstractSheet(const QString &name, int id, Workbook *workbook, AbstractSheetPrivate *d)
    : AbstractXmlFile(d)
{
	d_func()->name = name;
	d_func()->id = id;
	d_func()->workbook = workbook;
}

QString AbstractSheet::sheetName() const
{
	Q_D(const AbstractSheet);
	return d->name;
}

void AbstractSheet::setSheetName(const QString &sheetName)
{
	Q_D(AbstractSheet);
	d->name = sheetName;
}

AbstractSheet::Type AbstractSheet::sheetType() const
{
	Q_D(const AbstractSheet);
	return d->type;
}

void AbstractSheet::setSheetType(Type type)
{
	Q_D(AbstractSheet);
	d->type = type;
}

AbstractSheet::State AbstractSheet::sheetState() const
{
	Q_D(const AbstractSheet);
	return d->sheetState;
}

void AbstractSheet::setState(State state)
{
	Q_D(AbstractSheet);
	d->sheetState = state;
}

bool AbstractSheet::isHidden() const
{
	Q_D(const AbstractSheet);
	return d->sheetState != State::Visible;
}

bool AbstractSheet::isVisible() const
{
	return !isHidden();
}

void AbstractSheet::setHidden(bool hidden)
{
	Q_D(AbstractSheet);
	if (hidden == isHidden())
		return;

	d->sheetState = hidden ? State::Hidden : State::Visible;
}

void AbstractSheet::setVisible(bool visible)
{
	setHidden(!visible);
}

int AbstractSheet::sheetId() const
{
	Q_D(const AbstractSheet);
	return d->id;
}

Drawing *AbstractSheet::drawing() const
{
	Q_D(const AbstractSheet);
	return d->drawing.data();
}

Workbook *AbstractSheet::workbook() const
{
	Q_D(const AbstractSheet);
	return d->workbook;
}

QT_END_NAMESPACE_XLSX
