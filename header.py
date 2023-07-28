from pywebio.output import put_markdown, put_row, put_html,put_link



def header():
    put_row([
        put_markdown('# Kyeo个人工具箱'),
        put_link('首页','/'),
        put_link('建材月度销账数据处理','/jc')
    ], size='2fr auto').style('align-items:center')
    put_html('<script async defer src="https://buttons.github.io/buttons.js"></script>')

    # put_markdown(index_md_zh)

