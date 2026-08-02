# 4. 📈 تصميم الرسم البياني المطابق للمطلوب تماماً
    chart = ScatterChart()

    # وضع العنوان في المساحة الفارغة العلوية خارج منطقة الشبكة
    chart.title = str(test_title) if test_title else "B1"
    chart.title.overlay = False  # يضمن عدم تداخل العنوان مع الشبكة

    # إزالة أي لون خلفية للرسم ومساحة الشبكة لتظهر خلايا إكسل شفافة
    chart.graphicalProperties = GraphicalProperties()
    chart.graphicalProperties.noFill = True

    chart.plot_area.graphicalProperties = GraphicalProperties()
    chart.plot_area.graphicalProperties.noFill = True

    # إظهار خطوط الشبكة الأساسية الخفيفة
    chart.x_axis.majorGridlines = ChartLines()
    chart.y_axis.majorGridlines = ChartLines()

    # تنسيق الأرقام بمرتبتين عشريتين (0.00)
    chart.x_axis.number_format = "0.00"
    chart.y_axis.number_format = "0.00"

    xvalues = Reference(
        ws, min_col=2, min_row=start_cal_row, max_row=end_cal_row
    )
    yvalues = Reference(
        ws, min_col=3, min_row=start_cal_row, max_row=end_cal_row
    )

    series = Series(yvalues, xvalues, title_from_data=False)
    series.marker.symbol = "circle"  # نقاط دوائر زرقاء
    series.marker.size = 6
    series.graphicalProperties.line.noFill = True  # عدم وصل النقاط بخط صلب

    # خط الاتجاه المتقطع
    trendline = openpyxl.chart.trendline.Trendline(
        trendlineType="linear", dispEq=True, dispRSqr=True
    )
    try:
        trendline.graphicalProperties = GraphicalProperties()
        trendline.graphicalProperties.line = LineProperties(
            prstDash="sysDot", cmpd="sng"
        )
    except Exception:
        pass

    series.trendline = trendline

    chart.series.append(series)
    chart.legend = None
    chart.width = 13
    chart.height = 7.5
    ws.add_chart(chart, "F3")
