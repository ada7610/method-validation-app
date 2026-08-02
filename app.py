# 4. 📈 تصميم الرسم البياني المنسق حسب الطلب
    chart = ScatterChart()
    chart.title = str(test_title) if test_title else "B1"
    chart.title.overlay = False  # العنوان خارج شبكة الرسم

    # إزالة الخلفيات الملونة
    chart.graphicalProperties = GraphicalProperties()
    chart.graphicalProperties.noFill = True

    chart.plot_area.graphicalProperties = GraphicalProperties()
    chart.plot_area.graphicalProperties.noFill = True

    # خطوط الشبكة الأساسية
    chart.x_axis.majorGridlines = ChartLines()
    chart.y_axis.majorGridlines = ChartLines()

    # جعل لون خطوط محاور X و Y رمادي فاتح بدلاً من الأسود
    chart.x_axis.graphicalProperties = GraphicalProperties()
    chart.x_axis.graphicalProperties.line = LineProperties(solidFill="BFBFBF")

    chart.y_axis.graphicalProperties = GraphicalProperties()
    chart.y_axis.graphicalProperties.line = LineProperties(solidFill="BFBFBF")

    # جعل تنسيق الأرقام بـ 4 خانات عشرية (0.0000)
    chart.x_axis.number_format = "0.0000"
    chart.y_axis.number_format = "0.0000"

    xvalues = Reference(
        ws, min_col=2, min_row=start_cal_row, max_row=end_cal_row
    )
    yvalues = Reference(
        ws, min_col=3, min_row=start_cal_row, max_row=end_cal_row
    )

    series = Series(yvalues, xvalues, title_from_data=False)
    series.marker.symbol = "circle"
    series.marker.size = 6
    series.graphicalProperties.line.noFill = True  # نقاط بدون توصيل صلب

    # إضافة خط الاتجاه وإظهار معادلة Y و R2
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
