#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@文件        :excel_web
@说明        :第一个PyWebIO程序
@时间        :2023/07/17 09:45:55
@作者        :跳跃的🐸
@版本        :1.0
'''
from pywebio.pin import *
from pywebio.input import *
from pywebio.output import *
from pywebio.session import *
import pandas as pd
import re
from io import StringIO
import time
from header import header

def show_total(instance_id):
    re_str = '([京津沪渝冀豫云辽黑湘皖鲁新苏浙赣鄂桂甘晋蒙陕吉闽贵粤青藏川宁琼使领]{1}[A-Z]{1}(([A-Z0-9]{5}[DF]{1})|([DF]{1}[A-Z0-9]{5})))|([京津沪渝冀豫云辽黑湘皖鲁新苏浙赣鄂桂甘晋蒙陕吉闽贵粤青藏川宁琼使领]{1}[A-Z]{1}[A-Z0-9]{4}[A-Z0-9挂学警港澳]{1})'
    """return the current data in the datatable"""
    csv = eval_js("""
    window[`ag_grid_${instance_id}_promise`].then(function(gridOptions) {
        var csv = gridOptions.api.getDataAsCsv();
        return csv
    })""", instance_id=instance_id)
    df = pd.read_csv(StringIO(csv), float_precision='round_trip')
    # 根据提货车号内容判断正则判断是否车牌号，否则就是装船
    df['出库方式'] = df.提货车号.apply(lambda x: '装船' if re.findall(
        re_str, x.upper().replace('卾', '鄂')) == [] else '汽提')

    # 根据产地判断是否集港的方式
    df['入库方式'] = df.产地.apply(lambda x: '水运' if x == '九江萍钢' or x ==
                             '长钢' or x == '广钢' or x == '宁夏建龙(特钢)' else '汽运')

    df2 = pd.pivot_table(df, index=['入库方式', '产地'], columns=['出库方式'], values=[
                         '实发件数', '实发量'], aggfunc=['sum'], fill_value=0, margins_name='Total', margins=True)
    df3 = pd.concat([df2, df2.query('入库方式 != "Total"').sum(level=0).assign(
        产地='total').set_index('产地', append=True)]).sort_index()
    title = [
        [span("类别", row=2), span("产地", row=2),
         span("件数", col=3), span("重量", col=3)],
        ["汽提", "装船", "小计", "汽提", "装船", "小计"]
    ]
    data = df3.reset_index().values.tolist()
    out_in_car = []
    out_in_ship = []
    total_car = []
    total_ship = []
    total = []
    for row in data:
        for i in range(len(row)):
            if isinstance(row[i], (int, float)):
                row[i] = round(row[i], 3)
        if "Total" in row:
            row[0] = span("总计", col=2)
            del row[1]
            total.append(row)
        if "total" in row and "水运" in row:
            row[0] = span("水运合计", col=2)
            del row[1]
            total_ship.append(row)
        if "total" in row and "汽运" in row:
            row[0] = span("汽运合计", col=2)
            del row[1]
            total_car.append(row)
        if "total" not in row and "水运" in row:
            del row[0]
            out_in_ship.append(row)
        if "total" not in row and "汽运" in row:
            del row[0]
            out_in_car.append(row)
    out_in_ship[0].insert(0, span("水运", row=len(out_in_ship)))
    out_in_car[0].insert(0, span("汽运", row=len(out_in_car)))
    all_table = title + out_in_car + total_car + out_in_ship + total_ship + total
    with use_scope('total', clear=True):
        put_table(all_table)

def export_excel(instance_id):
    # 获取表格数据
    """return the current data in the datatable"""
    csv = eval_js("""
    window[`ag_grid_${instance_id}_promise`].then(function(gridOptions) {
        var csv = gridOptions.api.getDataAsCsv();
        return csv
    })""", instance_id=instance_id)
    df = pd.read_csv(StringIO(csv))
    # # 导出为Excel文件
    df.to_excel('table.xlsx', index=False)
    put_text('导出成功！')


def update(file):
    # 读取Excel文件
    df = pd.read_excel(file)
    # print(df.head())
    table = df.to_dict(orient='records')
    instance_id = "table"
    # 显示Excel文件内容
    with use_scope('table', clear=True):
        put_datatable(records=table, instance_id=instance_id,
                      theme='balham-dark', grid_args={
                          'pagination': 'true',
                          'paginationPageSize': '50',
                          'defaultColDef': {
                              'editable': 'true',
                              'resizable': 'false',
                              'wrapText': 'true',
                          },
                      })
        put_button(['统计汇总'], lambda: show_total(instance_id))
        put_button(['导出为Excel'], lambda: export_excel(instance_id))


def jc():
    header()
    put_file_upload(name='excel', placeholder='请选择要上传的Excel文件',
                            accept='.xlsx,.xls')
    while True:
        changed = pin_wait_change('excel')
        with use_scope('res', clear=True):
            put_table([['文件名', put_text(changed['value']['filename'])],
                      ['修改时间', put_markdown(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(changed['value']['last_modified'])))]])
            if changed['value'] is not None:
                if 'content' in changed['value']:
                    put_buttons(['显示文件内容'], lambda _: update(
                        changed['value']['content']))
            else:
                put_text("未选择任何文件，请选择文件！")
    # 通过文件上传控件上传Excel文件


# if __name__ == '__main__':
#     # 启动pywebio应用
#     start_server(jc, debug=True, port=8080)
