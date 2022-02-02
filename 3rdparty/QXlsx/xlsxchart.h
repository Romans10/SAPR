#pragma once

#include <QSharedPointer>
#include <QXmlStreamReader>
#include <QXmlStreamWriter>
#include <QtGlobal>

#include "xlsxabstractxmlfile.h"

QT_BEGIN_NAMESPACE_XLSX

class AbstractSheet;
class Worksheet;
class ChartPrivate;
class CellRange;
class DrawingAnchor;

class Chart : public AbstractXmlFile
{
	Q_DECLARE_PRIVATE(Chart)
public:
	enum class Type
	{  // 16 type of chart (ECMA 376)
		NoStatement = 0,  // Zero is internally used for unknown types
		Area,
		Area3D,
		Line,
		Line3D,
		Stock,
		Radar,
		Scatter,
		Pie,
		Pie3D,
		Doughnut,
		Bar,
		Bar3D,
		OfPie,
		Surface,
		Surface3D,
		Bubble,
	};

	enum class AxisPos
	{
		None = (-1),
		Left = 0,
		Right,
		Top,
		Bottom
	};

	~Chart();

	void addSeries(const CellRange &range, AbstractSheet *sheet = NULL, bool headerH = false, bool headerV = false, bool swapHeaders = false);
	void setChartType(Type type);
	void setChartStyle(int id);
	void setAxisTitle(Chart::AxisPos pos, QString axisTitle);
	void setChartTitle(QString strchartTitle);
	void setChartLegend(Chart::AxisPos legendPos, bool overlap = false);
	void setGridlinesEnable(bool majorGridlinesEnable = false, bool minorGridlinesEnable = false);

	bool loadFromXmlFile(QIODevice *device);
	void saveToXmlFile(QIODevice *device) const;

private:
	friend class AbstractSheet;
	friend class Worksheet;
	friend class Chartsheet;
	friend class DrawingAnchor;

	Chart(AbstractSheet *parent, CreateFlag flag);
};

QT_END_NAMESPACE_XLSX
