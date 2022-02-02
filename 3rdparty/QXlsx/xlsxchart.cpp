// xlsxchart.cpp

#include <QDebug>
#include <QIODevice>
#include <QString>
#include <QXmlStreamReader>
#include <QXmlStreamWriter>
#include <QtGlobal>

#include "xlsxcellrange.h"
#include "xlsxchart_p.h"
#include "xlsxutility_p.h"
#include "xlsxworksheet.h"

QT_BEGIN_NAMESPACE_XLSX

ChartPrivate::ChartPrivate(Chart *q, Chart::CreateFlag flag)
    : AbstractXmlFilePrivate(q, flag)
    , chartType(static_cast<Chart::Type>(0))
{
}

ChartPrivate::~ChartPrivate()
{
}

/*!
 * \internal
 */
Chart::Chart(AbstractSheet *parent, CreateFlag flag)
    : AbstractXmlFile(new ChartPrivate(this, flag))
{
	Q_D(Chart);

	d_func()->sheet = parent;

	// d->legendPos = Chart::AxisPos::None;
	d->legendPos = Chart::AxisPos::None;
	d->legendOverlay = false;
	d->majorGridlinesEnabled = false;
	d->minorGridlinesEnabled = false;
}

/*!
 * Destroys the chart.
 */
Chart::~Chart()
{
}

/*!
 * Add the data series which is in the range \a range of the \a sheet.
 */
void Chart::addSeries(const CellRange &range, AbstractSheet *sheet, bool headerH, bool headerV, bool swapHeaders)
{
	Q_D(Chart);

	if (!range.isValid())
		return;
	if (sheet && sheet->sheetType() != AbstractSheet::Type::Work)
		return;
	if (!sheet && d->sheet->sheetType() != AbstractSheet::Type::Work)
		return;

	QString sheetName = sheet ? sheet->sheetName() : d->sheet->sheetName();
	// In case sheetName contains space or '
	sheetName = escapeSheetName(sheetName);

	if (range.columnCount() == 1 || range.rowCount() == 1) {
		QSharedPointer<XlsxSeries> series = QSharedPointer<XlsxSeries>(new XlsxSeries);
		series->numberDataSource_numRef = sheetName + QLatin1String("!") + range.toString(true, true);
		d->seriesList.append(series);
	} else if ((range.columnCount() < range.rowCount()) || swapHeaders) {
		// Column based series
		int firstDataRow = range.firstRow();
		int firstDataColumn = range.firstColumn();

		QString axDataSouruce_numRef;
		if (d->chartType == Type::Scatter || d->chartType == Type::Bubble) {
			firstDataColumn += 1;
			CellRange subRange(range.firstRow(), range.firstColumn(), range.lastRow(), range.firstColumn());
			axDataSouruce_numRef = sheetName + QLatin1String("!") + subRange.toString(true, true);
		}

		if (headerH) {
			firstDataRow += 1;
		}
		if (headerV) {
			firstDataColumn += 1;
		}

		for (int col = firstDataColumn; col <= range.lastColumn(); ++col) {
			CellRange subRange(firstDataRow, col, range.lastRow(), col);
			QSharedPointer<XlsxSeries> series = QSharedPointer<XlsxSeries>(new XlsxSeries);
			series->axDataSource_numRef = axDataSouruce_numRef;
			series->numberDataSource_numRef = sheetName + QLatin1String("!") + subRange.toString(true, true);

			if (headerH) {
				CellRange subRange(range.firstRow(), col, range.firstRow(), col);
				series->headerH_numRef = sheetName + QLatin1String("!") + subRange.toString(true, true);
			} else {
				series->headerH_numRef = QString();
			}
			if (headerV) {
				CellRange subRange(firstDataRow, range.firstColumn(), range.lastRow(), range.firstColumn());
				series->headerV_numRef = sheetName + QLatin1String("!") + subRange.toString(true, true);
			} else {
				series->headerV_numRef = QString();
			}
			series->swapHeader = swapHeaders;

			d->seriesList.append(series);
		}

	} else {
		// Row based series
		int firstDataRow = range.firstRow();
		int firstDataColumn = range.firstColumn();

		QString axDataSouruce_numRef;
		if (d->chartType == Type::Scatter || d->chartType == Type::Bubble) {
			firstDataRow += 1;
			CellRange subRange(range.firstRow(), range.firstColumn(), range.firstRow(), range.lastColumn());
			axDataSouruce_numRef = sheetName + QLatin1String("!") + subRange.toString(true, true);
		}

		if (headerH) {
			firstDataRow += 1;
		}
		if (headerV) {
			firstDataColumn += 1;
		}

		for (int row = firstDataRow; row <= range.lastRow(); ++row) {
			CellRange subRange(row, firstDataColumn, row, range.lastColumn());
			QSharedPointer<XlsxSeries> series = QSharedPointer<XlsxSeries>(new XlsxSeries);
			series->axDataSource_numRef = axDataSouruce_numRef;
			series->numberDataSource_numRef = sheetName + QLatin1String("!") + subRange.toString(true, true);

			if (headerH) {
				CellRange subRange(range.firstRow(), firstDataColumn, range.firstRow(), range.lastColumn());
				series->headerH_numRef = sheetName + QLatin1String("!") + subRange.toString(true, true);
			} else {
				series->headerH_numRef = QString();
			}

			if (headerV) {
				CellRange subRange(row, range.firstColumn(), row, range.firstColumn());
				series->headerV_numRef = sheetName + QLatin1String("!") + subRange.toString(true, true);
			} else {
				series->headerV_numRef = QString();
			}
			series->swapHeader = swapHeaders;

			d->seriesList.append(series);
		}
	}
}

/*!
 * Set the type of the chart to \a type
 */
void Chart::setChartType(Type type)
{
	Q_D(Chart);

	d->chartType = type;
}

/*!
 * \internal
 *
 */
void Chart::setChartStyle(int id)
{
	Q_UNUSED(id)
	//! Todo
}

void Chart::setAxisTitle(Chart::AxisPos pos, QString axisTitle)
{
	Q_D(Chart);

	if (axisTitle.isEmpty())
		return;

	// dev24 : fixed for old compiler
	if (pos == Chart::AxisPos::Left) {
		d->axisNames[XlsxAxis::AxisPos::Left] = axisTitle;
	} else if (pos == Chart::AxisPos::Top) {
		d->axisNames[XlsxAxis::AxisPos::Top] = axisTitle;
	} else if (pos == Chart::AxisPos::Right) {
		d->axisNames[XlsxAxis::AxisPos::Right] = axisTitle;
	} else if (pos == Chart::AxisPos::Bottom) {
		d->axisNames[XlsxAxis::AxisPos::Bottom] = axisTitle;
	}
}

// dev25
void Chart::setChartTitle(QString strchartTitle)
{
	Q_D(Chart);

	d->chartTitle = strchartTitle;
}

void Chart::setChartLegend(Chart::AxisPos legendPos, bool overlay)
{
	Q_D(Chart);

	d->legendPos = legendPos;
	d->legendOverlay = overlay;
}

void Chart::setGridlinesEnable(bool majorGridlinesEnable, bool minorGridlinesEnable)
{
	Q_D(Chart);

	d->majorGridlinesEnabled = majorGridlinesEnable;
	d->minorGridlinesEnabled = minorGridlinesEnable;
}

/*!
 * \internal
 */
void Chart::saveToXmlFile(QIODevice *device) const
{
	Q_D(const Chart);

	/*
	    <chartSpace>
	        <chart>
	            <view3D>
	                <perspective val="30"/>
	            </view3D>
	            <plotArea>
	                <layout/>
	                <barChart>
	                ...
	                </barChart>
	                <catAx/>
	                <valAx/>
	            </plotArea>
	            <legend>
	            ...
	            </legend>
	        </chart>
	        <printSettings>
	        </printSettings>
	    </chartSpace>
	*/

	QXmlStreamWriter writer(device);

	writer.writeStartDocument(QStringLiteral("1.0"), true);

	// L.4.13.2.2 Chart
	//
	//  chartSpace is the root node, which contains an element defining the chart,
	// and an element defining the print settings for the chart.

	writer.writeStartElement(QStringLiteral("c:chartSpace"));

	writer.writeAttribute(QStringLiteral("xmlns:c"), QStringLiteral("http://schemas.openxmlformats.org/drawingml/2006/chart"));
	writer.writeAttribute(QStringLiteral("xmlns:a"), QStringLiteral("http://schemas.openxmlformats.org/drawingml/2006/main"));
	writer.writeAttribute(QStringLiteral("xmlns:r"), QStringLiteral("http://schemas.openxmlformats.org/officeDocument/2006/relationships"));

	/*
	 * chart is the root element for the chart. If the chart is a 3D chart,
	 * then a view3D element is contained, which specifies the 3D view.
	 * It then has a plot area, which defines a layout and contains an element
	 * that corresponds to, and defines, the type of chart.
	 */

	d->saveXmlChart(writer);

	writer.writeEndElement();  // c:chartSpace
	writer.writeEndDocument();
}

/*!
 * \internal
 */
bool Chart::loadFromXmlFile(QIODevice *device)
{
	Q_D(Chart);

	QXmlStreamReader reader(device);
	while (!reader.atEnd()) {
		reader.readNextStartElement();
		if (reader.tokenType() == QXmlStreamReader::StartElement) {
			if (reader.name() == QLatin1String("chart")) {
				if (!d->loadXmlChart(reader)) {
					return false;
				}
			}
		}
	}

	return true;
}

bool ChartPrivate::loadXmlChart(QXmlStreamReader &reader)
{
	Q_ASSERT(reader.name() == QLatin1String("chart"));

	//    qDebug() << "-------------- loadXmlChart";

	while (!reader.atEnd()) {
		reader.readNextStartElement();

		//        qDebug() << "-------------1- " << reader.name();

		if (reader.tokenType() == QXmlStreamReader::StartElement) {
			if (reader.name() == QLatin1String("plotArea")) {
				if (!loadXmlPlotArea(reader)) {
					return false;
				}
			} else if (reader.name() == QLatin1String("title")) {
				//! Todo

				if (loadXmlChartTitle(reader)) {
				}
			}
			//            else if (reader.name() == QLatin1String("legend"))
			//            {
			//                loadXmlChartLegend(reader);
			//                qDebug() << "-------------- loadXmlChartLegend";
			//            }
		} else if (reader.tokenType() == QXmlStreamReader::EndElement && reader.name() == QLatin1String("chart")) {
			break;
		}
	}
	return true;
}

// TO DEBUG: loop is not work, when i looping second element.
/*
dchrt_CT_PlotArea =
    element layout { dchrt_CT_Layout }?,
    (element areaChart { dchrt_Area }
        | element area3DChart { dchrt_ Area3D }
        | element lineChart { dchrt_Line }
        | element line3DChart { dchrt_Line3D }
        | element stockChart { dchrt_Stock }
        | element radarChart { dchrt_Radar }
        | element scatterChart { dchrt_Scatter }
        | element pieChart { dchrt_Pie }
        | element pie3DChart { dchrt_Pie3D }
        | element doughnutChart { dchrt_Doughnut }
        | element barChart { dchrt_Bar }
        | element bar3DChart { dchrt_Bar3D }
        | element ofPieChart { dchrt_OfPie }
        | element surfaceChart { dchrt_Surface }
        | element surface3DChart { dchrt_Surface3D }
        | element bubbleChart { dchrt_Bubble })+,
    (element valAx { dchrt_CT_ValAx }
        | element catAx { dchrt_CT_CatAx }
        | element dateAx { dchrt_CT_DateAx }
        | element serAx { dchrt_CT_SerAx })*,
    element dTable { dchrt_CT_DTable }?,
    element spPr { a_CT_ShapeProperties }?,
    element extLst { dchrt_CT_ExtensionList }?
 */
bool ChartPrivate::loadXmlPlotArea(QXmlStreamReader &reader)
{
	Q_ASSERT(reader.name() == QLatin1String("plotArea"));

	// TO DEBUG:

	reader.readNext();

	while (!reader.atEnd()) {
		//        qDebug() << "-------------2- " << reader.name();

		if (reader.isStartElement()) {
			if (!loadXmlPlotAreaElement(reader)) {
				qDebug() << "[debug] failed to load plotarea element.";
				return false;
			} else if (reader.name() == QLatin1String("legend"))  // Why here?
			{
				loadXmlChartLegend(reader);
				//                qDebug() << "-------------- loadXmlChartLegend";
			}

			reader.readNext();
		} else {
			reader.readNext();
		}
	}

	return true;
}

bool ChartPrivate::loadXmlPlotAreaElement(QXmlStreamReader &reader)
{
	if (reader.name() == QLatin1String("layout")) {
		//! ToDo extract attributes
		layout = readSubTree(reader);
	} else if (reader.name().endsWith(QLatin1String("Chart"))) {
		// for pieChart, barChart, ... (choose one)
		if (!loadXmlXxxChart(reader)) {
			qDebug() << "[debug] failed to load chart";
			return false;
		}
	} else if (reader.name() == QLatin1String("catAx"))  // choose one : catAx, dateAx, serAx, valAx
	{
		// qDebug() << "loadXmlAxisCatAx()";
		loadXmlAxisCatAx(reader);
	} else if (reader.name() == QLatin1String("dateAx"))  // choose one : catAx, dateAx, serAx, valAx
	{
		// qDebug() << "loadXmlAxisDateAx()";
		loadXmlAxisDateAx(reader);
	} else if (reader.name() == QLatin1String("serAx"))  // choose one : catAx, dateAx, serAx, valAx
	{
		// qDebug() << "loadXmlAxisSerAx()";
		loadXmlAxisSerAx(reader);
	} else if (reader.name() == QLatin1String("valAx"))  // choose one : catAx, dateAx, serAx, valAx
	{
		// qDebug() << "loadXmlAxisValAx()";
		loadXmlAxisValAx(reader);
	} else if (reader.name() == QLatin1String("dTable")) {
		//! ToDo
		// dTable "CT_DTable"
		// reader.skipCurrentElement();
	} else if (reader.name() == QLatin1String("spPr")) {
		//! ToDo
		// spPr "a:CT_ShapeProperties"
		// reader.skipCurrentElement();
	} else if (reader.name() == QLatin1String("extLst")) {
		//! ToDo
		// extLst "CT_ExtensionList"
		// reader.skipCurrentElement();
	}

	return true;
}

bool ChartPrivate::loadXmlXxxChart(QXmlStreamReader &reader)
{
	const auto &name = reader.name();

	if (name == QLatin1String("areaChart")) {
		chartType = Chart::Type::Area;
	} else if (name == QLatin1String("area3DChart")) {
		chartType = Chart::Type::Area3D;
	} else if (name == QLatin1String("lineChart")) {
		chartType = Chart::Type::Line;
	} else if (name == QLatin1String("line3DChart")) {
		chartType = Chart::Type::Line3D;
	} else if (name == QLatin1String("stockChart")) {
		chartType = Chart::Type::Stock;
	} else if (name == QLatin1String("radarChart")) {
		chartType = Chart::Type::Radar;
	} else if (name == QLatin1String("scatterChart")) {
		chartType = Chart::Type::Scatter;
	} else if (name == QLatin1String("pieChart")) {
		chartType = Chart::Type::Pie;
	} else if (name == QLatin1String("pie3DChart")) {
		chartType = Chart::Type::Pie3D;
	} else if (name == QLatin1String("doughnutChart")) {
		chartType = Chart::Type::Doughnut;
	} else if (name == QLatin1String("barChart")) {
		chartType = Chart::Type::Bar;
	} else if (name == QLatin1String("bar3DChart")) {
		chartType = Chart::Type::Bar3D;
	} else if (name == QLatin1String("ofPieChart")) {
		chartType = Chart::Type::OfPie;
	} else if (name == QLatin1String("surfaceChart")) {
		chartType = Chart::Type::Surface;
	} else if (name == QLatin1String("surface3DChart")) {
		chartType = Chart::Type::Surface3D;
	} else if (name == QLatin1String("bubbleChart")) {
		chartType = Chart::Type::Bubble;
	} else {
		qDebug() << "[undefined chart type] " << name;
		chartType = Chart::Type::NoStatement;
		return false;
	}

	while (!reader.atEnd()) {
		reader.readNextStartElement();
		if (reader.tokenType() == QXmlStreamReader::StartElement) {
			// dev57

			if (reader.name() == QLatin1String("ser")) {
				loadXmlSer(reader);
			} else if (reader.name() == QLatin1String("varyColors")) {
			} else if (reader.name() == QLatin1String("barDir")) {
			} else if (reader.name() == QLatin1String("axId")) {
				//

			} else if (reader.name() == QLatin1String("scatterStyle")) {
			} else if (reader.name() == QLatin1String("holeSize")) {
			} else {
			}

		} else if (reader.tokenType() == QXmlStreamReader::EndElement && reader.name() == name) {
			break;
		}
	}

	return true;
}

bool ChartPrivate::loadXmlSer(QXmlStreamReader &reader)
{
	Q_ASSERT(reader.name() == QLatin1String("ser"));

	QSharedPointer<XlsxSeries> series = QSharedPointer<XlsxSeries>(new XlsxSeries);
	seriesList.append(series);

	while (!reader.atEnd() && !(reader.tokenType() == QXmlStreamReader::EndElement && reader.name() == QLatin1String("ser"))) {
		if (reader.readNextStartElement()) {
			// TODO beide Header noch auswerten RTR 2019.11
			const auto &name = reader.name();
			if (name == QLatin1String("tx")) {
				while (!reader.atEnd() && !(reader.tokenType() == QXmlStreamReader::EndElement && reader.name() == name)) {
					if (reader.readNextStartElement()) {
						if (reader.name() == QLatin1String("strRef"))
							series->headerV_numRef = loadXmlStrRef(reader);
					}
				}
			} else if (name == QLatin1String("cat") || name == QLatin1String("xVal")) {
				while (!reader.atEnd() && !(reader.tokenType() == QXmlStreamReader::EndElement && reader.name() == name)) {
					if (reader.readNextStartElement()) {
						if (reader.name() == QLatin1String("numRef"))
							series->axDataSource_numRef = loadXmlNumRef(reader);
						else if (reader.name() == QLatin1String("strRef"))
							series->headerH_numRef = loadXmlStrRef(reader);
					}
				}
			} else if (name == QLatin1String("val") || name == QLatin1String("yVal")) {
				while (!reader.atEnd() && !(reader.tokenType() == QXmlStreamReader::EndElement && reader.name() == name)) {
					if (reader.readNextStartElement()) {
						if (reader.name() == QLatin1String("numRef"))
							series->numberDataSource_numRef = loadXmlNumRef(reader);
					}
				}
			} else if (name == QLatin1String("extLst")) {
				while (!reader.atEnd() && !(reader.tokenType() == QXmlStreamReader::EndElement && reader.name() == name)) {
					reader.readNextStartElement();
				}
			}
		}
	}

	return true;
}

QString ChartPrivate::loadXmlNumRef(QXmlStreamReader &reader)
{
	Q_ASSERT(reader.name() == QLatin1String("numRef"));

	while (!reader.atEnd() && !(reader.tokenType() == QXmlStreamReader::EndElement && reader.name() == QLatin1String("numRef"))) {
		if (reader.readNextStartElement()) {
			if (reader.name() == QLatin1String("f"))
				return reader.readElementText();
		}
	}

	return QString();
}

QString ChartPrivate::loadXmlStrRef(QXmlStreamReader &reader)
{
	Q_ASSERT(reader.name() == QLatin1String("strRef"));

	while (!reader.atEnd() && !(reader.tokenType() == QXmlStreamReader::EndElement && reader.name() == QLatin1String("strRef"))) {
		if (reader.readNextStartElement()) {
			if (reader.name() == QLatin1String("f"))
				return reader.readElementText();
		}
	}

	return QString();
}

void ChartPrivate::saveXmlChart(QXmlStreamWriter &writer) const
{
	//----------------------------------------------------
	// c:chart
	writer.writeStartElement(QStringLiteral("c:chart"));

	//----------------------------------------------------
	// c:title

	saveXmlChartTitle(writer);  // write 'chart title'

	//----------------------------------------------------
	// c:plotArea

	writer.writeStartElement(QStringLiteral("c:plotArea"));

	// a little workaround for Start- and EndElement with starting ">" and ending without ">"
	writer.device()->write("><c:layout>");  // layout
	writer.device()->write(layout.toUtf8());
	writer.device()->write("</c:layout");  // layout

	// dev35
	switch (chartType) {
	case Chart::Type::Area:
		saveXmlAreaChart(writer);
		break;
	case Chart::Type::Area3D:
		saveXmlAreaChart(writer);
		break;
	case Chart::Type::Line:
		saveXmlLineChart(writer);
		break;
	case Chart::Type::Line3D:
		saveXmlLineChart(writer);
		break;
	case Chart::Type::Stock:
		break;
	case Chart::Type::Radar:
		break;
	case Chart::Type::Scatter:
		saveXmlScatterChart(writer);
		break;
	case Chart::Type::Pie:
		saveXmlPieChart(writer);
		break;
	case Chart::Type::Pie3D:
		saveXmlPieChart(writer);
		break;
	case Chart::Type::Doughnut:
		saveXmlDoughnutChart(writer);
		break;
	case Chart::Type::Bar:
		saveXmlBarChart(writer);
		break;
	case Chart::Type::Bar3D:
		saveXmlBarChart(writer);
		break;
	case Chart::Type::OfPie:
		break;
	case Chart::Type::Surface:
		break;
	case Chart::Type::Surface3D:
		break;
	case Chart::Type::Bubble:
		break;
	default:
		break;
	}

	saveXmlAxis(writer);  // c:catAx, c:valAx, c:serAx, c:dateAx (choose one)

	//! TODO: write element
	// c:dTable CT_DTable
	// c:spPr   CT_ShapeProperties
	// c:extLst CT_ExtensionList

	writer.writeEndElement();  // c:plotArea

	// c:legend
	saveXmlChartLegend(writer);  // c:legend

	writer.writeEndElement();  // c:chart
}

bool ChartPrivate::loadXmlChartTitle(QXmlStreamReader &reader)
{
	//! TODO : load chart title

	Q_ASSERT(reader.name() == QLatin1String("title"));

	while (!reader.atEnd() && !(reader.tokenType() == QXmlStreamReader::EndElement && reader.name() == QLatin1String("title"))) {
		if (reader.readNextStartElement()) {
			if (reader.name() == QLatin1String("tx"))  // c:tx
				return loadXmlChartTitleTx(reader);
		}
	}

	return false;
}

bool ChartPrivate::loadXmlChartTitleTx(QXmlStreamReader &reader)
{
	Q_ASSERT(reader.name() == QLatin1String("tx"));

	while (!reader.atEnd() && !(reader.tokenType() == QXmlStreamReader::EndElement && reader.name() == QLatin1String("tx"))) {
		if (reader.readNextStartElement()) {
			if (reader.name() == QLatin1String("rich"))  // c:rich
				return loadXmlChartTitleTxRich(reader);
		}
	}

	return false;
}

bool ChartPrivate::loadXmlChartTitleTxRich(QXmlStreamReader &reader)
{
	Q_ASSERT(reader.name() == QLatin1String("rich"));

	while (!reader.atEnd() && !(reader.tokenType() == QXmlStreamReader::EndElement && reader.name() == QLatin1String("rich"))) {
		if (reader.readNextStartElement()) {
			if (reader.name() == QLatin1String("p"))  // a:p
				return loadXmlChartTitleTxRichP(reader);
		}
	}

	return false;
}

bool ChartPrivate::loadXmlChartTitleTxRichP(QXmlStreamReader &reader)
{
	Q_ASSERT(reader.name() == QLatin1String("p"));

	while (!reader.atEnd() && !(reader.tokenType() == QXmlStreamReader::EndElement && reader.name() == QLatin1String("p"))) {
		if (reader.readNextStartElement()) {
			if (reader.name() == QLatin1String("r"))  // a:r
				return loadXmlChartTitleTxRichP_R(reader);
		}
	}

	return false;
}

bool ChartPrivate::loadXmlChartTitleTxRichP_R(QXmlStreamReader &reader)
{
	Q_ASSERT(reader.name() == QLatin1String("r"));

	while (!reader.atEnd() && !(reader.tokenType() == QXmlStreamReader::EndElement && reader.name() == QLatin1String("r"))) {
		if (reader.readNextStartElement()) {
			if (reader.name() == QLatin1String("t"))  // a:t
			{
				QString textValue = reader.readElementText();
				this->chartTitle = textValue;
				return true;
			}
		}
	}

	return false;
}

// write 'chart title'
void ChartPrivate::saveXmlChartTitle(QXmlStreamWriter &writer) const
{
	if (chartTitle.isEmpty())
		return;

	writer.writeStartElement(QStringLiteral("c:title"));
	/*
	<xsd:complexType name="CT_Title">
	    <xsd:sequence>
	        <xsd:element name="tx"      type="CT_Tx"                minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="layout"  type="CT_Layout"            minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="overlay" type="CT_Boolean"           minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="spPr"    type="a:CT_ShapeProperties" minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="txPr"    type="a:CT_TextBody"        minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="extLst"  type="CT_ExtensionList"     minOccurs="0" maxOccurs="1"/>
	    </xsd:sequence>
	</xsd:complexType>
	*/

	writer.writeStartElement(QStringLiteral("c:tx"));
	/*
	<xsd:complexType name="CT_Tx">
	    <xsd:sequence>
	        <xsd:choice     minOccurs="1"   maxOccurs="1">
	        <xsd:element    name="strRef"   type="CT_StrRef"        minOccurs="1" maxOccurs="1"/>
	        <xsd:element    name="rich"     type="a:CT_TextBody"    minOccurs="1" maxOccurs="1"/>
	        </xsd:choice>
	    </xsd:sequence>
	</xsd:complexType>
	*/

	writer.writeStartElement(QStringLiteral("c:rich"));
	/*
	<xsd:complexType name="CT_TextBody">
	    <xsd:sequence>
	        <xsd:element name="bodyPr"      type="CT_TextBodyProperties"    minOccurs=" 1"  maxOccurs="1"/>
	        <xsd:element name="lstStyle"    type="CT_TextListStyle"         minOccurs="0"   maxOccurs="1"/>
	        <xsd:element name="p"           type="CT_TextParagraph"         minOccurs="1"   maxOccurs="unbounded"/>
	    </xsd:sequence>
	</xsd:complexType>
	*/

	writer.writeEmptyElement(QStringLiteral("a:bodyPr"));  // <a:bodyPr/>
	/*
	<xsd:complexType name="CT_TextBodyProperties">
	    <xsd:sequence>
	        <xsd:element name="prstTxWarp" type="CT_PresetTextShape" minOccurs="0" maxOccurs="1"/>
	        <xsd:group ref="EG_TextAutofit" minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="scene3d" type="CT_Scene3D" minOccurs="0" maxOccurs="1"/>
	        <xsd:group ref="EG_Text3D" minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
	    </xsd:sequence>
	    <xsd:attribute name="rot" type="ST_Angle" use="optional"/>
	    <xsd:attribute name="spcFirstLastPara" type="xsd:boolean" use="optional"/>
	    <xsd:attribute name="vertOverflow" type="ST_TextVertOverflowType" use="optional"/>
	    <xsd:attribute name="horzOverflow" type="ST_TextHorzOverflowType" use="optional"/>
	    <xsd:attribute name="vert" type="ST_TextVerticalType" use="optional"/>
	    <xsd:attribute name="wrap" type="ST_TextWrappingType" use="optional"/>
	    <xsd:attribute name="lIns" type="ST_Coordinate32" use="optional"/>
	    <xsd:attribute name="tIns" type="ST_Coordinate32" use="optional"/>
	    <xsd:attribute name="rIns" type="ST_Coordinate32" use="optional"/>
	    <xsd:attribute name="bIns" type="ST_Coordinate32" use="optional"/>
	    <xsd:attribute name="numCol" type="ST_TextColumnCount" use="optional"/>
	    <xsd:attribute name="spcCol" type="ST_PositiveCoordinate32" use="optional"/>
	    <xsd:attribute name="rtlCol" type="xsd:boolean" use="optional"/>
	    <xsd:attribute name="fromWordArt" type="xsd:boolean" use="optional"/>
	    <xsd:attribute name="anchor" type="ST_TextAnchoringType" use="optional"/>
	    <xsd:attribute name="anchorCtr" type="xsd:boolean" use="optional"/>
	    <xsd:attribute name="forceAA" type="xsd:boolean" use="optional"/>
	    <xsd:attribute name="upright" type="xsd:boolean" use="optional" default="false"/>
	    <xsd:attribute name="compatLnSpc" type="xsd:boolean" use="optional"/>
	</xsd:complexType>
	 */

	writer.writeEmptyElement(QStringLiteral("a:lstStyle"));  // <a:lstStyle/>

	writer.writeStartElement(QStringLiteral("a:p"));
	/*
	<xsd:complexType name="CT_TextParagraph">
	    <xsd:sequence>
	        <xsd:element    name="pPr"          type="CT_TextParagraphProperties" minOccurs="0" maxOccurs="1"/>
	        <xsd:group      ref="EG_TextRun"    minOccurs="0" maxOccurs="unbounded"/>
	        <xsd:element    name="endParaRPr"   type="CT_TextCharacterProperties" minOccurs="0"
	        maxOccurs="1"/>
	    </xsd:sequence>
	</xsd:complexType>
	 */

	// <a:pPr lvl="0">
	writer.writeStartElement(QStringLiteral("a:pPr"));

	writer.writeAttribute(QStringLiteral("lvl"), QStringLiteral("0"));

	// <a:defRPr b="0"/>
	writer.writeStartElement(QStringLiteral("a:defRPr"));

	writer.writeAttribute(QStringLiteral("b"), QStringLiteral("0"));

	writer.writeEndElement();  // a:defRPr

	writer.writeEndElement();  // a:pPr

	/*
	<xsd:group name="EG_TextRun">
	    <xsd:choice>
	        <xsd:element name="r"   type="CT_RegularTextRun"/>
	        <xsd:element name="br"  type="CT_TextLineBreak"/>
	        <xsd:element name="fld" type="CT_TextField"/>
	    </xsd:choice>
	</xsd:group>
	*/

	writer.writeStartElement(QStringLiteral("a:r"));
	/*
	<xsd:complexType name="CT_RegularTextRun">
	    <xsd:sequence>
	        <xsd:element name="rPr" type="CT_TextCharacterProperties" minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="t"   type="xsd:string" minOccurs="1" maxOccurs="1"/>
	    </xsd:sequence>
	</xsd:complexType>
	 */

	// <a:t>chart name</a:t>
	writer.writeTextElement(QStringLiteral("a:t"), chartTitle);

	writer.writeEndElement();  // a:r

	writer.writeEndElement();  // a:p

	writer.writeEndElement();  // c:rich

	writer.writeEndElement();  // c:tx

	// <c:overlay val="0"/>
	writer.writeStartElement(QStringLiteral("c:overlay"));
	writer.writeAttribute(QStringLiteral("val"), QStringLiteral("0"));
	writer.writeEndElement();  // c:overlay

	writer.writeEndElement();  // c:title
}
// }}

// write 'chart legend'
void ChartPrivate::saveXmlChartLegend(QXmlStreamWriter &writer) const
{
	if (legendPos == Chart::AxisPos::None)
		return;

	//    <c:legend>
	//    <c:legendPos val="r"/>
	//    <c:overlay val="0"/>
	//    </c:legend>

	writer.writeStartElement(QStringLiteral("c:legend"));

	writer.writeStartElement(QStringLiteral("c:legendPos"));

	QString pos;
	switch (legendPos) {
	// case Chart::AxisPos::Right:
	case Chart::AxisPos::Right:
		pos = QStringLiteral("r");
		break;

	// case Chart::AxisPos::Left:
	case Chart::AxisPos::Left:
		pos = QStringLiteral("l");
		break;

	// case Chart::AxisPos::Top:
	case Chart::AxisPos::Top:
		pos = QStringLiteral("t");
		break;

	// case Chart::AxisPos::Bottom:
	case Chart::AxisPos::Bottom:
		pos = QStringLiteral("b");
		break;

	default:
		pos = QStringLiteral("r");
		break;
	}

	writer.writeAttribute(QStringLiteral("val"), pos);

	writer.writeEndElement();  // c:legendPos

	writer.writeStartElement(QStringLiteral("c:overlay"));

	if (legendOverlay) {
		writer.writeAttribute(QStringLiteral("val"), QStringLiteral("1"));
	} else {
		writer.writeAttribute(QStringLiteral("val"), QStringLiteral("0"));
	}

	writer.writeEndElement();  // c:overlay

	writer.writeEndElement();  // c:legend
}

void ChartPrivate::saveXmlPieChart(QXmlStreamWriter &writer) const
{
	QString name = chartType == Chart::Type::Pie ? QStringLiteral("c:pieChart") : QStringLiteral("c:pie3DChart");

	writer.writeStartElement(name);

	// Do the same behavior as Excel, Pie prefer varyColors
	writer.writeEmptyElement(QStringLiteral("c:varyColors"));
	writer.writeAttribute(QStringLiteral("val"), QStringLiteral("1"));

	for (int i = 0; i < seriesList.size(); ++i)
		saveXmlSer(writer, seriesList[i].data(), i);

	writer.writeEndElement();  // pieChart, pie3DChart
}

void ChartPrivate::saveXmlBarChart(QXmlStreamWriter &writer) const
{
	QString name = chartType == Chart::Type::Bar ? QStringLiteral("c:barChart") : QStringLiteral("c:bar3DChart");

	writer.writeStartElement(name);

	writer.writeEmptyElement(QStringLiteral("c:barDir"));
	writer.writeAttribute(QStringLiteral("val"), QStringLiteral("col"));

	for (int i = 0; i < seriesList.size(); ++i) {
		saveXmlSer(writer, seriesList[i].data(), i);
	}

	if (axisList.isEmpty()) {
		const_cast<ChartPrivate *>(this)->axisList.append(
		    QSharedPointer<XlsxAxis>(new XlsxAxis(XlsxAxis::Type::Cat, XlsxAxis::AxisPos::Bottom, 0, 1, axisNames[XlsxAxis::AxisPos::Bottom])));

		const_cast<ChartPrivate *>(this)->axisList.append(
		    QSharedPointer<XlsxAxis>(new XlsxAxis(XlsxAxis::Type::Val, XlsxAxis::AxisPos::Left, 1, 0, axisNames[XlsxAxis::AxisPos::Left])));
	}

	// Note: Bar3D have 2~3 axes
	// int axisListSize = axisList.size();
	// [dev62]
	// Q_ASSERT( axisListSize == 2 ||
	//          ( axisListSize == 3 && chartType == Chart::Type::Bar3D ) );

	for (int i = 0; i < axisList.size(); ++i) {
		writer.writeEmptyElement(QStringLiteral("c:axId"));
		writer.writeAttribute(QStringLiteral("val"), QString::number(axisList[i]->axisId));
	}

	writer.writeEndElement();  // barChart, bar3DChart
}

void ChartPrivate::saveXmlLineChart(QXmlStreamWriter &writer) const
{
	QString name = chartType == Chart::Type::Line ? QStringLiteral("c:lineChart") : QStringLiteral("c:line3DChart");

	writer.writeStartElement(name);

	// writer.writeEmptyElement(QStringLiteral("grouping")); // dev22

	for (int i = 0; i < seriesList.size(); ++i)
		saveXmlSer(writer, seriesList[i].data(), i);

	if (axisList.isEmpty()) {
		const_cast<ChartPrivate *>(this)->axisList.append(
		    QSharedPointer<XlsxAxis>(new XlsxAxis(XlsxAxis::Type::Cat, XlsxAxis::AxisPos::Bottom, 0, 1, axisNames[XlsxAxis::AxisPos::Bottom])));
		const_cast<ChartPrivate *>(this)->axisList.append(
		    QSharedPointer<XlsxAxis>(new XlsxAxis(XlsxAxis::Type::Val, XlsxAxis::AxisPos::Left, 1, 0, axisNames[XlsxAxis::AxisPos::Left])));
		if (chartType == Chart::Type::Line3D)
			const_cast<ChartPrivate *>(this)->axisList.append(QSharedPointer<XlsxAxis>(new XlsxAxis(XlsxAxis::Type::Ser, XlsxAxis::AxisPos::Bottom, 2, 0)));
	}

	Q_ASSERT((axisList.size() == 2 || chartType == Chart::Type::Line) || (axisList.size() == 3 && chartType == Chart::Type::Line3D));

	for (int i = 0; i < axisList.size(); ++i) {
		writer.writeEmptyElement(QStringLiteral("c:axId"));
		writer.writeAttribute(QStringLiteral("val"), QString::number(axisList[i]->axisId));
	}

	writer.writeEndElement();  // lineChart, line3DChart
}

void ChartPrivate::saveXmlScatterChart(QXmlStreamWriter &writer) const
{
	const QString name = QStringLiteral("c:scatterChart");

	writer.writeStartElement(name);

	writer.writeEmptyElement(QStringLiteral("c:scatterStyle"));

	for (int i = 0; i < seriesList.size(); ++i)
		saveXmlSer(writer, seriesList[i].data(), i);

	if (axisList.isEmpty()) {
		const_cast<ChartPrivate *>(this)->axisList.append(
		    QSharedPointer<XlsxAxis>(new XlsxAxis(XlsxAxis::Type::Val, XlsxAxis::AxisPos::Bottom, 0, 1, axisNames[XlsxAxis::AxisPos::Bottom])));
		const_cast<ChartPrivate *>(this)->axisList.append(
		    QSharedPointer<XlsxAxis>(new XlsxAxis(XlsxAxis::Type::Val, XlsxAxis::AxisPos::Left, 1, 0, axisNames[XlsxAxis::AxisPos::Left])));
	}

	int axisListSize = axisList.size();
	Q_ASSERT(axisListSize == 2);

	for (int i = 0; i < axisList.size(); ++i) {
		writer.writeEmptyElement(QStringLiteral("c:axId"));
		writer.writeAttribute(QStringLiteral("val"), QString::number(axisList[i]->axisId));
	}

	writer.writeEndElement();  // c:scatterChart
}

void ChartPrivate::saveXmlAreaChart(QXmlStreamWriter &writer) const
{
	QString name = chartType == Chart::Type::Area ? QStringLiteral("c:areaChart") : QStringLiteral("c:area3DChart");

	writer.writeStartElement(name);

	// writer.writeEmptyElement(QStringLiteral("grouping")); // dev22

	for (int i = 0; i < seriesList.size(); ++i)
		saveXmlSer(writer, seriesList[i].data(), i);

	if (axisList.isEmpty()) {
		const_cast<ChartPrivate *>(this)->axisList.append(QSharedPointer<XlsxAxis>(new XlsxAxis(XlsxAxis::Type::Cat, XlsxAxis::AxisPos::Bottom, 0, 1)));
		const_cast<ChartPrivate *>(this)->axisList.append(QSharedPointer<XlsxAxis>(new XlsxAxis(XlsxAxis::Type::Val, XlsxAxis::AxisPos::Left, 1, 0)));
	}

	// Note: Area3D have 2~3 axes
	Q_ASSERT(axisList.size() == 2 || (axisList.size() == 3 && chartType == Chart::Type::Area3D));

	for (int i = 0; i < axisList.size(); ++i) {
		writer.writeEmptyElement(QStringLiteral("c:axId"));
		writer.writeAttribute(QStringLiteral("val"), QString::number(axisList[i]->axisId));
	}

	writer.writeEndElement();  // lineChart, line3DChart
}

void ChartPrivate::saveXmlDoughnutChart(QXmlStreamWriter &writer) const
{
	QString name = QStringLiteral("c:doughnutChart");

	writer.writeStartElement(name);

	writer.writeEmptyElement(QStringLiteral("c:varyColors"));
	writer.writeAttribute(QStringLiteral("val"), QStringLiteral("1"));

	for (int i = 0; i < seriesList.size(); ++i)
		saveXmlSer(writer, seriesList[i].data(), i);

	writer.writeStartElement(QStringLiteral("c:holeSize"));
	writer.writeAttribute(QStringLiteral("val"), QString::number(50));

	writer.writeEndElement();
}

void ChartPrivate::saveXmlSer(QXmlStreamWriter &writer, XlsxSeries *ser, int id) const
{
	writer.writeStartElement(QStringLiteral("c:ser"));
	writer.writeEmptyElement(QStringLiteral("c:idx"));
	writer.writeAttribute(QStringLiteral("val"), QString::number(id));
	writer.writeEmptyElement(QStringLiteral("c:order"));
	writer.writeAttribute(QStringLiteral("val"), QString::number(id));

	QString header1;
	QString header2;
	if (ser->swapHeader) {
		header1 = ser->headerH_numRef;
		header2 = ser->headerV_numRef;
	} else {
		header1 = ser->headerV_numRef;
		header2 = ser->headerH_numRef;
	}

	if (!header1.isEmpty()) {
		writer.writeStartElement(QStringLiteral("c:tx"));
		writer.writeStartElement(QStringLiteral("c:strRef"));
		writer.writeTextElement(QStringLiteral("c:f"), header1);
		writer.writeEndElement();
		writer.writeEndElement();
	}
	if (!header2.isEmpty()) {
		writer.writeStartElement(QStringLiteral("c:cat"));
		writer.writeStartElement(QStringLiteral("c:strRef"));
		writer.writeTextElement(QStringLiteral("c:f"), header2);
		writer.writeEndElement();
		writer.writeEndElement();
	}

#if 0
    if (!ser->axDataSource_numRef.isEmpty())
    {
        if (chartType == Chart::Type::Scatter || chartType == Chart::Type::Bubble)
        {
            writer.writeStartElement(QStringLiteral("c:xVal"));
        }
        else
        {
            writer.writeStartElement(QStringLiteral("c:cat"));
        }

        writer.writeStartElement(QStringLiteral("c:numRef"));
        writer.writeTextElement(QStringLiteral("c:f"), ser->axDataSource_numRef);
        writer.writeEndElement();//c:numRef
        writer.writeEndElement();//c:cat or c:xVal
    }
#endif

	if (!ser->numberDataSource_numRef.isEmpty()) {
		if (chartType == Chart::Type::Scatter || chartType == Chart::Type::Bubble)
			writer.writeStartElement(QStringLiteral("c:yVal"));
		else
			writer.writeStartElement(QStringLiteral("c:val"));
		writer.writeStartElement(QStringLiteral("c:numRef"));
		writer.writeTextElement(QStringLiteral("c:f"), ser->numberDataSource_numRef);
		writer.writeEndElement();  // c:numRef
		writer.writeEndElement();  // c:val or c:yVal
	}

	writer.writeEndElement();  // c:ser
}

bool ChartPrivate::loadXmlAxisCatAx(QXmlStreamReader &reader)
{
	XlsxAxis *axis = new XlsxAxis();
	axis->type = XlsxAxis::Type::Cat;
	axisList.append(QSharedPointer<XlsxAxis>(axis));

	// load EG_AxShared
	if (!loadXmlAxisEG_AxShared(reader, axis)) {
		qDebug() << "failed to load EG_AxShared";
		return false;
	}

	//! TODO: load element
	// auto
	// lblAlgn
	// lblOffset
	// tickLblSkip
	// tickMarkSkip
	// noMultiLvlLbl
	// extLst

	return true;
}

bool ChartPrivate::loadXmlAxisDateAx(QXmlStreamReader &reader)
{
	XlsxAxis *axis = new XlsxAxis();
	axis->type = XlsxAxis::Type::Date;
	axisList.append(QSharedPointer<XlsxAxis>(axis));

	// load EG_AxShared
	if (!loadXmlAxisEG_AxShared(reader, axis)) {
		qDebug() << "failed to load EG_AxShared";
		return false;
	}

	//! TODO: load element
	// auto
	// lblOffset
	// baseTimeUnit
	// majorUnit
	// majorTimeUnit
	// minorUnit
	// minorTimeUnit
	// extLst

	return true;
}

bool ChartPrivate::loadXmlAxisSerAx(QXmlStreamReader &reader)
{
	XlsxAxis *axis = new XlsxAxis();
	axis->type = XlsxAxis::Type::Ser;
	axisList.append(QSharedPointer<XlsxAxis>(axis));

	// load EG_AxShared
	if (!loadXmlAxisEG_AxShared(reader, axis)) {
		qDebug() << "failed to load EG_AxShared";
		return false;
	}

	//! TODO: load element
	// tickLblSkip
	// tickMarkSkip
	// extLst

	return true;
}

bool ChartPrivate::loadXmlAxisValAx(QXmlStreamReader &reader)
{
	Q_ASSERT(reader.name() == QLatin1String("valAx"));

	XlsxAxis *axis = new XlsxAxis();
	axis->type = XlsxAxis::Type::Val;
	axisList.append(QSharedPointer<XlsxAxis>(axis));

	if (!loadXmlAxisEG_AxShared(reader, axis)) {
		qDebug() << "failed to load EG_AxShared";
		return false;
	}

	//! TODO: load element
	// crossBetween
	// majorUnit
	// minorUnit
	// dispUnits
	// extLst

	return true;
}

/*
<xsd:group name="EG_AxShared">
    <xsd:sequence>
        <xsd:element name="axId" type="CT_UnsignedInt" minOccurs="1" maxOccurs="1"/> (*)(M)
        <xsd:element name="scaling" type="CT_Scaling" minOccurs="1" maxOccurs="1"/> (*)(M)
        <xsd:element name="delete" type="CT_Boolean" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="axPos" type="CT_AxPos" minOccurs="1" maxOccurs="1"/> (*)(M)
        <xsd:element name="majorGridlines" type="CT_ChartLines" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="minorGridlines" type="CT_ChartLines" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="title" type="CT_Title" minOccurs="0" maxOccurs="1"/> (*)
        <xsd:element name="numFmt" type="CT_NumFmt" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="majorTickMark" type="CT_TickMark" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="minorTickMark" type="CT_TickMark" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="tickLblPos" type="CT_TickLblPos" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="spPr" type="a:CT_ShapeProperties" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="txPr" type="a:CT_TextBody" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="crossAx" type="CT_UnsignedInt" minOccurs="1" maxOccurs="1"/> (*)(M)
        <xsd:choice minOccurs="0" maxOccurs="1">
            <xsd:element name="crosses" type="CT_Crosses" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="crossesAt" type="CT_Double" minOccurs="1" maxOccurs="1"/>
        </xsd:choice>
    </xsd:sequence>
</xsd:group>
*/
bool ChartPrivate::loadXmlAxisEG_AxShared(QXmlStreamReader &reader, XlsxAxis *axis)
{
	Q_ASSERT(NULL != axis);
	Q_ASSERT(reader.name().endsWith(QLatin1String("Ax")));
	QString name = reader.name().toString();  //

	while (!reader.atEnd()) {
		reader.readNextStartElement();
		if (reader.tokenType() == QXmlStreamReader::StartElement) {
			// qDebug() << "[debug]" << QTime::currentTime() << reader.name().toString();

			if (reader.name() == QLatin1String("axId"))  // mandatory element
			{
				// dev57
				uint axId = reader.attributes().value(QStringLiteral("val")).toString().toUInt();  // for Qt5.1
				axis->axisId = axId;
			} else if (reader.name() == QLatin1String("scaling")) {
				// mandatory element

				loadXmlAxisEG_AxShared_Scaling(reader, axis);
			} else if (reader.name() == QLatin1String("delete")) {
				//! TODO
			} else if (reader.name() == QLatin1String("axPos")) {
				// mandatory element

				QString axPosVal = reader.attributes().value(QLatin1String("val")).toString();

				if (axPosVal == QLatin1String("l")) {
					axis->axisPos = XlsxAxis::AxisPos::Left;
				} else if (axPosVal == QLatin1String("r")) {
					axis->axisPos = XlsxAxis::AxisPos::Right;
				} else if (axPosVal == QLatin1String("t")) {
					axis->axisPos = XlsxAxis::AxisPos::Top;
				} else if (axPosVal == QLatin1String("b")) {
					axis->axisPos = XlsxAxis::AxisPos::Bottom;
				}
			} else if (reader.name() == QLatin1String("majorGridlines")) {
				//! TODO anything else?
				majorGridlinesEnabled = true;
			} else if (reader.name() == QLatin1String("minorGridlines")) {
				//! TODO anything else?
				minorGridlinesEnabled = true;
			} else if (reader.name() == QLatin1String("title")) {
				// title
				if (!loadXmlAxisEG_AxShared_Title(reader, axis)) {
					qDebug() << "failed to load EG_AxShared title.";
					Q_ASSERT(false);
					return false;
				}
			} else if (reader.name() == QLatin1String("numFmt")) {
				//! TODO
			} else if (reader.name() == QLatin1String("majorTickMark")) {
				//! TODO
			} else if (reader.name() == QLatin1String("minorTickMark")) {
				//! TODO
			} else if (reader.name() == QLatin1String("tickLblPos")) {
				//! TODO
			} else if (reader.name() == QLatin1String("spPr")) {
				//! TODO
			} else if (reader.name() == QLatin1String("txPr")) {
				//! TODO
			} else if (reader.name() == QLatin1String("crossAx"))  // mandatory element
			{
				// dev57
				uint crossAx = reader.attributes().value(QLatin1String("val")).toString().toUInt();  // for Qt5.1
				axis->crossAx = crossAx;
			} else if (reader.name() == QLatin1String("crosses")) {
				//! TODO
			} else if (reader.name() == QLatin1String("crossesAt")) {
				//! TODO
			}

			// reader.readNext();
		} else if (reader.tokenType() == QXmlStreamReader::EndElement && reader.name().toString() == name) {
			break;
		}
	}

	return true;
}

bool ChartPrivate::loadXmlAxisEG_AxShared_Scaling(QXmlStreamReader &reader, XlsxAxis *axis)
{
	Q_UNUSED(axis);
	Q_ASSERT(reader.name() == QLatin1String("scaling"));

	while (!reader.atEnd()) {
		reader.readNextStartElement();
		if (reader.tokenType() == QXmlStreamReader::StartElement) {
			if (reader.name() == QLatin1String("orientation")) {
			} else {
			}
		} else if (reader.tokenType() == QXmlStreamReader::EndElement && reader.name() == QLatin1String("scaling")) {
			break;
		}
	}

	return true;
}

/*
  <xsd:complexType name="CT_Title">
      <xsd:sequence>
          <xsd:element name="tx" type="CT_Tx" minOccurs="0" maxOccurs="1"/>
          <xsd:element name="layout" type="CT_Layout" minOccurs="0" maxOccurs="1"/>
          <xsd:element name="overlay" type="CT_Boolean" minOccurs="0" maxOccurs="1"/>
          <xsd:element name="spPr" type="a:CT_ShapeProperties" minOccurs="0" maxOccurs="1"/>
          <xsd:element name="txPr" type="a:CT_TextBody" minOccurs="0" maxOccurs="1"/>
          <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
      </xsd:sequence>
  </xsd:complexType>

<xsd:complexType name="CT_Tx">
    <xsd:sequence>
        <xsd:choice minOccurs="1" maxOccurs="1">
            <xsd:element name="strRef" type="CT_StrRef" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="rich" type="a:CT_TextBody" minOccurs="1" maxOccurs="1"/>
        </xsd:choice>
    </xsd:sequence>
</xsd:complexType>

<xsd:complexType name="CT_StrRef">
    <xsd:sequence>
        <xsd:element name="f" type="xsd:string" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="strCache" type="CT_StrData" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
</xsd:complexType>

<xsd:complexType name="CT_TextBody">
    <xsd:sequence>
        <xsd:element name="bodyPr" type="CT_TextBodyProperties" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="lstStyle" type="CT_TextListStyle" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="p" type="CT_TextParagraph" minOccurs="1" maxOccurs="unbounded"/>
    </xsd:sequence>
</xsd:complexType>
  */
bool ChartPrivate::loadXmlAxisEG_AxShared_Title(QXmlStreamReader &reader, XlsxAxis *axis)
{
	Q_ASSERT(reader.name() == QLatin1String("title"));

	while (!reader.atEnd()) {
		reader.readNextStartElement();
		if (reader.tokenType() == QXmlStreamReader::StartElement) {
			if (reader.name() == QLatin1String("tx")) {
				loadXmlAxisEG_AxShared_Title_Tx(reader, axis);
			} else if (reader.name() == QLatin1String("overlay")) {
				//! TODO: load overlay
				loadXmlAxisEG_AxShared_Title_Overlay(reader, axis);
			} else {
			}
		} else if (reader.tokenType() == QXmlStreamReader::EndElement && reader.name() == QLatin1String("title")) {
			break;
		}
	}

	return true;
}

bool ChartPrivate::loadXmlAxisEG_AxShared_Title_Overlay(QXmlStreamReader &reader, XlsxAxis *axis)
{
	Q_UNUSED(axis);
	Q_ASSERT(reader.name() == QLatin1String("overlay"));

	while (!reader.atEnd()) {
		reader.readNextStartElement();
		if (reader.tokenType() == QXmlStreamReader::StartElement) {
		} else if (reader.tokenType() == QXmlStreamReader::EndElement && reader.name() == QLatin1String("overlay")) {
			break;
		}
	}

	return true;
}

bool ChartPrivate::loadXmlAxisEG_AxShared_Title_Tx(QXmlStreamReader &reader, XlsxAxis *axis)
{
	Q_ASSERT(reader.name() == QLatin1String("tx"));

	while (!reader.atEnd()) {
		reader.readNextStartElement();
		if (reader.tokenType() == QXmlStreamReader::StartElement) {
			if (reader.name() == QLatin1String("rich")) {
				loadXmlAxisEG_AxShared_Title_Tx_Rich(reader, axis);
			} else {
			}
		} else if (reader.tokenType() == QXmlStreamReader::EndElement && reader.name() == QLatin1String("tx")) {
			break;
		}
	}

	return true;
}

bool ChartPrivate::loadXmlAxisEG_AxShared_Title_Tx_Rich(QXmlStreamReader &reader, XlsxAxis *axis)
{
	Q_ASSERT(reader.name() == QLatin1String("rich"));

	while (!reader.atEnd()) {
		reader.readNextStartElement();
		if (reader.tokenType() == QXmlStreamReader::StartElement) {
			if (reader.name() == QLatin1String("p")) {
				loadXmlAxisEG_AxShared_Title_Tx_Rich_P(reader, axis);
			} else {
			}
		} else if (reader.tokenType() == QXmlStreamReader::EndElement && reader.name() == QLatin1String("rich")) {
			break;
		}
	}

	return true;
}

bool ChartPrivate::loadXmlAxisEG_AxShared_Title_Tx_Rich_P(QXmlStreamReader &reader, XlsxAxis *axis)
{
	Q_ASSERT(reader.name() == QLatin1String("p"));

	while (!reader.atEnd()) {
		reader.readNextStartElement();
		if (reader.tokenType() == QXmlStreamReader::StartElement) {
			if (reader.name() == QLatin1String("r")) {
				loadXmlAxisEG_AxShared_Title_Tx_Rich_P_R(reader, axis);
			} else if (reader.name() == QLatin1String("pPr")) {
				loadXmlAxisEG_AxShared_Title_Tx_Rich_P_pPr(reader, axis);
			} else {
			}
		} else if (reader.tokenType() == QXmlStreamReader::EndElement && reader.name() == QLatin1String("p")) {
			break;
		}
	}

	return true;
}

bool ChartPrivate::loadXmlAxisEG_AxShared_Title_Tx_Rich_P_pPr(QXmlStreamReader &reader, XlsxAxis *axis)
{
	Q_UNUSED(axis);
	Q_ASSERT(reader.name() == QLatin1String("pPr"));

	while (!reader.atEnd()) {
		reader.readNextStartElement();
		if (reader.tokenType() == QXmlStreamReader::StartElement) {
			if (reader.name() == QLatin1String("defRPr")) {
				reader.readElementText();
			} else {
			}
		} else if (reader.tokenType() == QXmlStreamReader::EndElement && reader.name() == QLatin1String("pPr")) {
			break;
		}
	}

	return true;
}

bool ChartPrivate::loadXmlAxisEG_AxShared_Title_Tx_Rich_P_R(QXmlStreamReader &reader, XlsxAxis *axis)
{
	Q_ASSERT(reader.name() == QLatin1String("r"));

	while (!reader.atEnd()) {
		reader.readNextStartElement();
		if (reader.tokenType() == QXmlStreamReader::StartElement) {
			if (reader.name() == QLatin1String("t")) {
				QString strAxisName = reader.readElementText();
				XlsxAxis::AxisPos axisPos = axis->axisPos;
				axis->axisNames[axisPos] = strAxisName;
			} else {
			}
		} else if (reader.tokenType() == QXmlStreamReader::EndElement && reader.name() == QLatin1String("r")) {
			break;
		}
	}

	return true;
}

/*
<xsd:complexType name="CT_PlotArea">
    <xsd:sequence>
        <xsd:element name="layout" type="CT_Layout" minOccurs="0" maxOccurs="1"/>
        <xsd:choice minOccurs="1" maxOccurs="unbounded">
            <xsd:element name="areaChart" type="Area" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="area3DChart" type="Area3D" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="lineChart" type="Line" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="line3DChart" type="Line3D" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="stockChart" type="Stock" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="radarChart" type="Radar" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="scatterChart" type="Scatter" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="pieChart" type="Pie" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="pie3DChart" type="Pie3D" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="doughnutChart" type="Doughnut" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="barChart" type="Bar" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="bar3DChart" type="Bar3D" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="ofPieChart" type="OfPie" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="surfaceChart" type="Surface" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="surface3DChart" type="Surface3D" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="bubbleChart" type="Bubble" minOccurs="1" maxOccurs="1"/>
        </xsd:choice>
        <xsd:choice minOccurs="0" maxOccurs="unbounded">
            <xsd:element name="valAx" type="CT_ValAx" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="catAx" type="CT_CatAx" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="dateAx" type="CT_DateAx" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="serAx" type="CT_SerAx" minOccurs="1" maxOccurs="1"/>
        </xsd:choice>
        <xsd:element name="dTable" type="CT_DTable" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="spPr" type="a:CT_ShapeProperties" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
</xsd:complexType>
*/

/*
<xsd:complexType name="CT_CatAx">
    <xsd:sequence>
        <xsd:group ref="EG_AxShared" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="auto" type="CT_Boolean" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="lblAlgn" type="CT_LblAlgn" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="lblOffset" type="CT_LblOffset" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="tickLblSkip" type="CT_Skip" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="tickMarkSkip" type="CT_Skip" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="noMultiLvlLbl" type="CT_Boolean" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
</xsd:complexType>
<!----------------------------------------------------------------------------->
<xsd:complexType name="CT_DateAx">
    <xsd:sequence>
        <xsd:group ref="EG_AxShared" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="auto" type="CT_Boolean" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="lblOffset" type="CT_LblOffset" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="baseTimeUnit" type="CT_TimeUnit" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="majorUnit" type="CT_AxisUnit" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="majorTimeUnit" type="CT_TimeUnit" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="minorUnit" type="CT_AxisUnit" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="minorTimeUnit" type="CT_TimeUnit" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
</xsd:complexType>
<!----------------------------------------------------------------------------->
<xsd:complexType name="CT_SerAx">
    <xsd:sequence>
    <xsd:group ref="EG_AxShared" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="tickLblSkip" type="CT_Skip" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="tickMarkSkip" type="CT_Skip" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
</xsd:complexType>
<!----------------------------------------------------------------------------->
<xsd:complexType name="CT_ValAx">
    <xsd:sequence>
        <xsd:group ref="EG_AxShared" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="crossBetween" type="CT_CrossBetween" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="majorUnit" type="CT_AxisUnit" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="minorUnit" type="CT_AxisUnit" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="dispUnits" type="CT_DispUnits" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
</xsd:complexType>
*/

void ChartPrivate::saveXmlAxis(QXmlStreamWriter &writer) const
{
	for (int i = 0; i < axisList.size(); ++i) {
		XlsxAxis *axis = axisList[i].data();
		if (NULL == axis)
			continue;

		if (axis->type == XlsxAxis::Type::Cat) {
			saveXmlAxisCatAx(writer, axis);
		}
		if (axis->type == XlsxAxis::Type::Val) {
			saveXmlAxisValAx(writer, axis);
		}
		if (axis->type == XlsxAxis::Type::Ser) {
			saveXmlAxisSerAx(writer, axis);
		}
		if (axis->type == XlsxAxis::Type::Date) {
			saveXmlAxisDateAx(writer, axis);
		}
	}
}

void ChartPrivate::saveXmlAxisCatAx(QXmlStreamWriter &writer, XlsxAxis *axis) const
{
	/*
	<xsd:complexType name="CT_CatAx">
	    <xsd:sequence>
	        <xsd:group ref="EG_AxShared" minOccurs="1" maxOccurs="1"/>
	        <xsd:element name="auto" type="CT_Boolean" minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="lblAlgn" type="CT_LblAlgn" minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="lblOffset" type="CT_LblOffset" minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="tickLblSkip" type="CT_Skip" minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="tickMarkSkip" type="CT_Skip" minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="noMultiLvlLbl" type="CT_Boolean" minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
	    </xsd:sequence>
	</xsd:complexType>
	*/

	writer.writeStartElement(QStringLiteral("c:catAx"));

	saveXmlAxisEG_AxShared(writer, axis);  // EG_AxShared

	//! TODO: write element
	// auto
	// lblAlgn
	// lblOffset
	// tickLblSkip
	// tickMarkSkip
	// noMultiLvlLbl
	// extLst

	writer.writeEndElement();  // c:catAx
}

void ChartPrivate::saveXmlAxisDateAx(QXmlStreamWriter &writer, XlsxAxis *axis) const
{
	/*
	<xsd:complexType name="CT_DateAx">
	    <xsd:sequence>
	        <xsd:group ref="EG_AxShared" minOccurs="1" maxOccurs="1"/>
	        <xsd:element name="auto" type="CT_Boolean" minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="lblOffset" type="CT_LblOffset" minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="baseTimeUnit" type="CT_TimeUnit" minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="majorUnit" type="CT_AxisUnit" minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="majorTimeUnit" type="CT_TimeUnit" minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="minorUnit" type="CT_AxisUnit" minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="minorTimeUnit" type="CT_TimeUnit" minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
	    </xsd:sequence>
	</xsd:complexType>
	*/

	writer.writeStartElement(QStringLiteral("c:dateAx"));

	saveXmlAxisEG_AxShared(writer, axis);  // EG_AxShared

	//! TODO: write element
	// auto
	// lblOffset
	// baseTimeUnit
	// majorUnit
	// majorTimeUnit
	// minorUnit
	// minorTimeUnit
	// extLst

	writer.writeEndElement();  // c:dateAx
}

void ChartPrivate::saveXmlAxisSerAx(QXmlStreamWriter &writer, XlsxAxis *axis) const
{
	/*
	<xsd:complexType name="CT_SerAx">
	    <xsd:sequence>
	    <xsd:group ref="EG_AxShared" minOccurs="1" maxOccurs="1"/>
	        <xsd:element name="tickLblSkip" type="CT_Skip" minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="tickMarkSkip" type="CT_Skip" minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
	    </xsd:sequence>
	</xsd:complexType>
	*/

	writer.writeStartElement(QStringLiteral("c:serAx"));

	saveXmlAxisEG_AxShared(writer, axis);  // EG_AxShared

	//! TODO: write element
	// tickLblSkip
	// tickMarkSkip
	// extLst

	writer.writeEndElement();  // c:serAx
}

void ChartPrivate::saveXmlAxisValAx(QXmlStreamWriter &writer, XlsxAxis *axis) const
{
	/*
	<xsd:complexType name="CT_ValAx">
	    <xsd:sequence>
	        <xsd:group ref="EG_AxShared" minOccurs="1" maxOccurs="1"/>
	        <xsd:element name="crossBetween" type="CT_CrossBetween" minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="majorUnit" type="CT_AxisUnit" minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="minorUnit" type="CT_AxisUnit" minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="dispUnits" type="CT_DispUnits" minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
	    </xsd:sequence>
	</xsd:complexType>
	*/

	writer.writeStartElement(QStringLiteral("c:valAx"));

	saveXmlAxisEG_AxShared(writer, axis);  // EG_AxShared

	//! TODO: write element
	// crossBetween
	// majorUnit
	// minorUnit
	// dispUnits
	// extLst

	writer.writeEndElement();  // c:valAx
}

void ChartPrivate::saveXmlAxisEG_AxShared(QXmlStreamWriter &writer, XlsxAxis *axis) const
{
	/*
	<xsd:group name="EG_AxShared">
	    <xsd:sequence>
	        <xsd:element name="axId" type="CT_UnsignedInt" minOccurs="1" maxOccurs="1"/> (*)
	        <xsd:element name="scaling" type="CT_Scaling" minOccurs="1" maxOccurs="1"/> (*)
	        <xsd:element name="delete" type="CT_Boolean" minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="axPos" type="CT_AxPos" minOccurs="1" maxOccurs="1"/> (*)
	        <xsd:element name="majorGridlines" type="CT_ChartLines" minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="minorGridlines" type="CT_ChartLines" minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="title" type="CT_Title" minOccurs="0" maxOccurs="1"/> (***********************)
	        <xsd:element name="numFmt" type="CT_NumFmt" minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="majorTickMark" type="CT_TickMark" minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="minorTickMark" type="CT_TickMark" minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="tickLblPos" type="CT_TickLblPos" minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="spPr" type="a:CT_ShapeProperties" minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="txPr" type="a:CT_TextBody" minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="crossAx" type="CT_UnsignedInt" minOccurs="1" maxOccurs="1"/> (*)
	        <xsd:choice minOccurs="0" maxOccurs="1">
	            <xsd:element name="crosses" type="CT_Crosses" minOccurs="1" maxOccurs="1"/>
	            <xsd:element name="crossesAt" type="CT_Double" minOccurs="1" maxOccurs="1"/>
	        </xsd:choice>
	    </xsd:sequence>
	</xsd:group>
	*/

	writer.writeEmptyElement(QStringLiteral("c:axId"));  // 21.2.2.9. axId (Axis ID) (mandatory value)
	writer.writeAttribute(QStringLiteral("val"), QString::number(axis->axisId));

	writer.writeStartElement(QStringLiteral("c:scaling"));  // CT_Scaling (mandatory value)
	writer.writeEmptyElement(QStringLiteral("c:orientation"));  // CT_Orientation
	writer.writeAttribute(QStringLiteral("val"), QStringLiteral("minMax"));  // ST_Orientation
	writer.writeEndElement();  // c:scaling

	writer.writeEmptyElement(QStringLiteral("c:axPos"));  // axPos CT_AxPos (mandatory value)
	QString pos = GetAxisPosString(axis->axisPos);
	if (!pos.isEmpty()) {
		writer.writeAttribute(QStringLiteral("val"), pos);  // ST_AxPos
	}

	if (majorGridlinesEnabled) {
		writer.writeEmptyElement(QStringLiteral("c:majorGridlines"));
	}
	if (minorGridlinesEnabled) {
		writer.writeEmptyElement(QStringLiteral("c:minorGridlines"));
	}

	saveXmlAxisEG_AxShared_Title(writer, axis);  // "c:title" CT_Title

	writer.writeEmptyElement(QStringLiteral("c:crossAx"));  // crossAx (mandatory value)
	writer.writeAttribute(QStringLiteral("val"), QString::number(axis->crossAx));
}

void ChartPrivate::saveXmlAxisEG_AxShared_Title(QXmlStreamWriter &writer, XlsxAxis *axis) const
{
	// CT_Title

	/*
	<xsd:complexType name="CT_Title">
	    <xsd:sequence>
	        <xsd:element name="tx" type="CT_Tx" minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="layout" type="CT_Layout" minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="overlay" type="CT_Boolean" minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="spPr" type="a:CT_ShapeProperties" minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="txPr" type="a:CT_TextBody" minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
	    </xsd:sequence>
	</xsd:complexType>
	*/
	/*
	<xsd:complexType name="CT_Tx">
	    <xsd:sequence>
	        <xsd:choice minOccurs="1" maxOccurs="1">
	            <xsd:element name="strRef" type="CT_StrRef" minOccurs="1" maxOccurs="1"/>
	            <xsd:element name="rich" type="a:CT_TextBody" minOccurs="1" maxOccurs="1"/>
	        </xsd:choice>
	    </xsd:sequence>
	</xsd:complexType>
	*/
	/*
	<xsd:complexType name="CT_StrRef">
	    <xsd:sequence>
	        <xsd:element name="f" type="xsd:string" minOccurs="1" maxOccurs="1"/>
	        <xsd:element name="strCache" type="CT_StrData" minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
	    </xsd:sequence>
	</xsd:complexType>
	*/
	/*
	<xsd:complexType name="CT_TextBody">
	    <xsd:sequence>
	        <xsd:element name="bodyPr" type="CT_TextBodyProperties" minOccurs="1" maxOccurs="1"/>
	        <xsd:element name="lstStyle" type="CT_TextListStyle" minOccurs="0" maxOccurs="1"/>
	        <xsd:element name="p" type="CT_TextParagraph" minOccurs="1" maxOccurs="unbounded"/>
	    </xsd:sequence>
	</xsd:complexType>
	*/

	writer.writeStartElement(QStringLiteral("c:title"));

	// CT_Tx {{
	writer.writeStartElement(QStringLiteral("c:tx"));

	writer.writeStartElement(QStringLiteral("c:rich"));  // CT_TextBody

	writer.writeEmptyElement(QStringLiteral("a:bodyPr"));  // CT_TextBodyProperties

	writer.writeEmptyElement(QStringLiteral("a:lstStyle"));  // CT_TextListStyle

	writer.writeStartElement(QStringLiteral("a:p"));

	writer.writeStartElement(QStringLiteral("a:pPr"));
	writer.writeAttribute(QStringLiteral("lvl"), QString::number(0));

	writer.writeStartElement(QStringLiteral("a:defRPr"));
	writer.writeAttribute(QStringLiteral("b"), QString::number(0));
	writer.writeEndElement();  // a:defRPr
	writer.writeEndElement();  // a:pPr

	writer.writeStartElement(QStringLiteral("a:r"));
	QString strAxisName = GetAxisName(axis);
	writer.writeTextElement(QStringLiteral("a:t"), strAxisName);
	writer.writeEndElement();  // a:r

	writer.writeEndElement();  // a:p

	writer.writeEndElement();  // c:rich

	writer.writeEndElement();  // c:tx
	// CT_Tx }}

	writer.writeStartElement(QStringLiteral("c:overlay"));
	writer.writeAttribute(QStringLiteral("val"), QString::number(0));  // CT_Boolean
	writer.writeEndElement();  // c:overlay

	writer.writeEndElement();  // c:title
}

QString ChartPrivate::GetAxisPosString(XlsxAxis::AxisPos axisPos) const
{
	QString pos;
	switch (axisPos) {
	case XlsxAxis::AxisPos::Top:
		pos = QStringLiteral("t");
		break;
	case XlsxAxis::AxisPos::Bottom:
		pos = QStringLiteral("b");
		break;
	case XlsxAxis::AxisPos::Left:
		pos = QStringLiteral("l");
		break;
	case XlsxAxis::AxisPos::Right:
		pos = QStringLiteral("r");
		break;
	default:
		break;  // ??
	}

	return pos;
}

QString ChartPrivate::GetAxisName(XlsxAxis *axis) const
{
	QString strAxisName;
	if (NULL == axis)
		return strAxisName;

	QString pos = GetAxisPosString(axis->axisPos);  // l, t, r, b
	if (pos.isEmpty())
		return strAxisName;

	strAxisName = axis->axisNames[axis->axisPos];
	return strAxisName;
}

///
/// \brief ChartPrivate::readSubTree
/// \param reader
/// \return
///
QString ChartPrivate::readSubTree(QXmlStreamReader &reader)
{
	QString treeString;
	QString prefix;
	const auto &treeName = reader.name();

	while (!reader.atEnd()) {
		reader.readNextStartElement();
		if (reader.tokenType() == QXmlStreamReader::StartElement) {
			prefix = reader.prefix().toString();

			treeString += QLatin1String("<") + reader.qualifiedName().toString();

			const QXmlStreamAttributes attributes = reader.attributes();
			for (const QXmlStreamAttribute &attr : attributes) {
				treeString += QLatin1String(" ") + attr.name().toString() + QLatin1String("=\"") + attr.value().toString() + QLatin1String("\"");
			}
			treeString += QStringLiteral(">");
		} else if (reader.tokenType() == QXmlStreamReader::EndElement) {
			if (reader.name() == treeName) {
				break;
			}
			treeString += QLatin1String("</") + reader.qualifiedName().toString() + QLatin1String(">");
		}
	}

	return treeString;
}

///
/// \brief ChartPrivate::loadXmlChartLegend
/// \param reader
/// \return
///
bool ChartPrivate::loadXmlChartLegend(QXmlStreamReader &reader)
{
	Q_ASSERT(reader.name() == QLatin1String("legend"));

	while (!reader.atEnd() && !(reader.tokenType() == QXmlStreamReader::EndElement && reader.name() == QLatin1String("legend"))) {
		if (reader.readNextStartElement()) {
			if (reader.name() == QLatin1String("legendPos"))  // c:legendPos
			{
				QString pos = reader.attributes().value(QLatin1String("val")).toString();
				if (pos.compare(QLatin1String("r"), Qt::CaseInsensitive) == 0) {
					// legendPos = Chart::AxisPos::Right;
					legendPos = Chart::AxisPos::Right;
				} else if (pos.compare(QLatin1String("l"), Qt::CaseInsensitive) == 0) {
					// legendPos = Chart::AxisPos::Left;
					legendPos = Chart::AxisPos::Left;
				} else if (pos.compare(QLatin1String("t"), Qt::CaseInsensitive) == 0) {
					// legendPos = Chart::AxisPos::Top;
					legendPos = Chart::AxisPos::Top;
				} else if (pos.compare(QLatin1String("b"), Qt::CaseInsensitive) == 0) {
					// legendPos = Chart::AxisPos::Bottom;
					legendPos = Chart::AxisPos::Bottom;
				} else {
					// legendPos = Chart::AxisPos::None;
					legendPos = Chart::AxisPos::None;
				}
			} else if (reader.name() == QLatin1String("overlay"))  // c:legendPos
			{
				QString pos = reader.attributes().value(QLatin1String("val")).toString();
				if (pos.compare(QLatin1String("1"), Qt::CaseInsensitive) == 0) {
					legendOverlay = true;
				} else {
					legendOverlay = false;
				}
			}
		}
	}

	return false;
}

QT_END_NAMESPACE_XLSX
