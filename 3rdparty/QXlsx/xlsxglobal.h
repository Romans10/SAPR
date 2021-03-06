#pragma once

#include <cstdio>
#include <iostream>
#include <string>

#include <QtGlobal>

#define QT_BEGIN_NAMESPACE_XLSX namespace QXlsx {
#define QT_END_NAMESPACE_XLSX }

#define QXLSX_USE_NAMESPACE using namespace QXlsx;

#if QT_VERSION < QT_VERSION_CHECK(5, 0, 0)
#	undef QT_BEGIN_MOC_NAMESPACE
#	undef QT_END_MOC_NAMESPACE
#	define QT_BEGIN_MOC_NAMESPACE namespace QXlsx {
#	define QT_END_MOC_NAMESPACE }
#	define QStringLiteral QLatin1String
#	define Q_DECL_NOTHROW
#endif
