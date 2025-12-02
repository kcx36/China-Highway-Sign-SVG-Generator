import os
import pandas as pd

# 配置变量
EXCEL_PATH = r"input.xlsx" # Excel输入
SAVE_DIR = r"D:\test" # 保存路径
INKSCAPE_PATH = r"C:\Program Files\Inkscape\bin\inkscape.exe" # Inkscape路径
NAMING_STYLE = 2 # 命名方式 (1=中文命名, 2=英文命名)

# 省份英文对照表
PROVINCE_ENGLISH = {
    "京": "Beijing", "津": "Tianjin", "冀": "Hebei", "晋": "Shanxi",
    "蒙": "Inner Mongolia", "辽": "Liaoning", "吉": "Jilin",
    "黑": "Heilongjiang", "沪": "Shanghai", "苏": "Jiangsu",
    "浙": "Zhejiang", "皖": "Anhui", "闽": "Fujian", "赣": "Jiangxi",
    "鲁": "Shandong", "豫": "Henan", "鄂": "Hubei", "湘": "Hunan",
    "粤": "Guangdong", "桂": "Guangxi", "琼": "Hainan", "渝": "Chongqing",
    "川": "Sichuan", "黔": "Guizhou", "滇": "Yunnan", "藏": "Tibet",
    "陕": "Shaanxi", "甘": "Gansu", "青": "Qinghai", "宁": "Ningxia",
    "新": "Xinjiang"
}

# 模板定义
TEMPLATES = {
    # 国家高速模板
    "national_1digit_noname": '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1000 1000" width="500" height="500" xmlns:v="https://vecta.io/nano"><path d="M880 1000H120C53.7 1000 0 946.3 0 880V120C0 53.7 53.7 0 120 0h760c66.3 0 120 53.7 120 120v760c0 66.3-53.7 120-120 120z" fill="#fff"/><path d="M880 970H120c-49.7 0-90-40.3-90-90V230h940v650c0 49.7-40.3 90-90 90z" fill="#08963b"/><path d="M970 230H30V120c0-49.7 40.3-90 90-90h760c49.7 0 90 40.3 90 90z" fill="#e71f20"/><text xml:space="preserve" x="650.965" y="812.778" font-weight="bold" font-size="722.315" font-family="Source Han Sans" letter-spacing="0" fill="#fbf9f9"><tspan x="500" y="812.778" font-weight="normal" font-family="B型交通标志专用字体" text-anchor="middle">HIGHWAY_NUMBER</tspan></text><text x="500" y="160" letter-spacing="100" fill="#fbf9f9" font-size="100" font-family="A型交通标志专用字体"><tspan x="550" y="160" text-anchor="middle">国家高速</tspan></text></svg>''',

    "national_2digit_noname": '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1250 1000" width="625" height="500" xmlns:v="https://vecta.io/nano"><path d="M1130 1000H120C53.7 1000 0 946.3 0 880V120C0 53.7 53.7 0 120 0h1010c66.3 0 120 53.7 120 120v760c0 66.3-53.7 120-120 120z" fill="#fff"/><path d="M1130 970H120c-49.7 0-90-40.3-90-90V230h1190v650c0 49.7-40.3 90-90 90z" fill="#08963b"/><path d="M1220 230H30V120c0-49.7 40.3-90 90-90h1010c49.7 0 90 40.3 90 90z" fill="#e71f20"/><text xml:space="preserve" x="776.695" y="812.743" font-weight="bold" font-size="725.81" font-family="Source Han Sans" letter-spacing="0" fill="#fbf9f9"><tspan x="625" y="812.743" font-weight="normal" font-family="B型交通标志专用字体" text-anchor="middle">HIGHWAY_NUMBER</tspan></text><text x="625" y="160" letter-spacing="100" fill="#fbf9f9" font-size="100" font-family="A型交通标志专用字体"><tspan x="675" y="160" text-anchor="middle">国家高速</tspan></text></svg>''',

    "national_4digit_noname": '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1700 1000" width="850" height="500" xmlns:v="https://vecta.io/nano"><path d="M1580 1000H120C53.7 1000 0 946.3 0 880V120C0 53.7 53.7 0 120 0h1460c66.3 0 120 53.7 120 120v760c0 66.3-53.7 120-120 120z" fill="#fff"/><path d="M1575 970H125c-49.7 0-90-40.3-90-90V230h1630v650c0 49.7-40.3 90-90 90z" fill="#08963b"/><path d="M1575 30H125c-49.7 0-90 40.3-90 90v110h1630V120c0-49.7-40.3-90-90-90z" fill="#e71f20"/><text xml:space="preserve" x="1001.207" y="812.766" font-weight="bold" font-size="723.475" font-family="Source Han Sans" letter-spacing="0" fill="#fbf9f9"><tspan x="850" y="812.766" font-weight="normal" font-family="B型交通标志专用字体" text-anchor="middle">HIGHWAY_NUMBER_PART1<tspan font-size="475.799">HIGHWAY_NUMBER_PART2</tspan></tspan></text><text x="850" y="160" letter-spacing="200" fill="#fbf9f9" font-size="100" font-family="A型交通标志专用字体"><tspan text-anchor="middle">国家高速</tspan></text></svg>''',

    "national_2digit_4char": '''<svg xmlns="http://www.w3.org/2000/svg" width="625" height="600" viewBox="0 0 1250 1200" xmlns:v="https://vecta.io/nano"><path d="M1130 1200H120c-66.3 0-120-53.7-120-120V120C0 53.7 53.7 0 120 0h1010c66.3 0 120 53.7 120 120v960c0 66.3-53.7 120-120 120z" fill="#08963b"/><path d="M1130 1170H120c-49.7 0-90-40.3-90-90V120c0-49.7 40.3-90 90-90h1010c49.7 0 90 40.3 90 90v960c0 49.7-40.3 90-90 90z" fill="#fff"/><path d="M1130 1140H120c-33.1 0-60-26.9-60-60V260h1130v820c0 33.1-26.9 60-60 60z" fill="#08963b"/><path d="M1190 260H60V120c0-33.1 26.9-60 60-60h1010c33.1 0 60 26.9 60 60z" fill="#e71f20"/><text xml:space="preserve" x="776.695" y="782.743" font-size="725.81" letter-spacing="0" font-weight="bold" font-family="Source Han Sans" fill="#fbf9f9"><tspan x="625" y="782.743" font-weight="normal" font-family="B型交通标志专用字体" text-anchor="middle">HIGHWAY_NUMBER</tspan></text><text x="624.8" y="1020" letter-spacing="50" fill="#fbf9f9" font-family="A型交通标志专用字体" font-size="200"><tspan text-anchor="middle">HIGHWAY_NAME</tspan></text><text x="625" y="190" letter-spacing="100" fill="#fbf9f9" font-family="A型交通标志专用字体" font-size="100"><tspan text-anchor="middle">国家高速</tspan></text></svg>''',

    "national_2digit_5char": '''<svg xmlns="http://www.w3.org/2000/svg" width="625" height="600" viewBox="0 0 1250 1200" xmlns:v="https://vecta.io/nano"><path d="M1130 1200H120c-66.3 0-120-53.7-120-120V120C0 53.7 53.7 0 120 0h1010c66.3 0 120 53.7 120 120v960c0 66.3-53.7 120-120 120z" fill="#08963b"/><path d="M1130 1170H120c-49.7 0-90-40.3-90-90V120c0-49.7 40.3-90 90-90h1010c49.7 0 90 40.3 90 90v960c0 49.7-40.3 90-90 90z" fill="#fff"/><path d="M1130 1140H120c-33.1 0-60-26.9-60-60V260h1130v820c0 33.1-26.9 60-60 60z" fill="#08963b"/><path d="M1190 260H60V120c0-33.1 26.9-60 60-60h1010c33.1 0 60 26.9 60 60z" fill="#e71f20"/><text xml:space="preserve" x="776.695" y="782.743" font-size="725.81" letter-spacing="0" font-weight="bold" font-family="Source Han Sans" fill="#fbf9f9"><tspan x="625" y="782.743" font-weight="normal" font-family="B型交通标志专用字体" text-anchor="middle">HIGHWAY_NUMBER</tspan></text><text x="624.8" y="1020" letter-spacing="0" fill="#fbf9f9" font-family="A型交通标志专用字体" font-size="200"><tspan letter-spacing="0" text-anchor="middle">HIGHWAY_NAME</tspan></text><text x="625" y="190" letter-spacing="100" fill="#fbf9f9" font-family="A型交通标志专用字体" font-size="100"><tspan text-anchor="middle">国家高速</tspan></text></svg>''',

    "national_2digit_6char": '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1250 1200" width="625" height="600" xmlns:v="https://vecta.io/nano"><path d="M1130 1200H120c-66.3 0-120-53.7-120-120V120C0 53.7 53.7 0 120 0h1010c66.3 0 120 53.7 120 120v960c0 66.3-53.7 120-120 120z" fill="#08963b"/><path d="M1130 1170H120c-49.7 0-90-40.3-90-90V120c0-49.7 40.3-90 90-90h1010c49.7 0 90 40.3 90 90v960c0 49.7-40.3 90-90 90z" fill="#fff"/><path d="M1130 1140H120c-33.1 0-60-26.9-60-60V260h1130v820c0 33.1-26.9 60-60 60z" fill="#08963b"/><path d="M1190 260H60V120c0-33.1 26.9-60 60-60h1010c33.1 0 60 26.9 60 60z" fill="#e71f20"/><text xml:space="preserve" x="1050.871" y="786.794" font-weight="bold" font-size="740" font-family="Source Han Sans" letter-spacing="0" fill="#fbf9f9"><tspan x="625" y="786.794" font-weight="normal" font-family="B型交通标志专用字体" text-anchor="middle">HIGHWAY_NUMBER</tspan></text><text x="625" y="1007.874" letter-spacing="10" fill="#fbf9f9" font-family="A型交通标志专用字体" font-size="160"><tspan text-anchor="middle">HIGHWAY_NAME</tspan></text><text x="625" y="190" letter-spacing="100" fill="#fbf9f9" font-family="A型交通标志专用字体" font-size="100"><tspan text-anchor="middle">国家高速</tspan></text></svg>''',
    
    "national_2digit_7char": '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1250 1200" width="625" height="600" xmlns:v="https://vecta.io/nano"><path d="M1130 1200H120c-66.3 0-120-53.7-120-120V120C0 53.7 53.7 0 120 0h1010c66.3 0 120 53.7 120 120v960c0 66.3-53.7 120-120 120z" fill="#08963b"/><path d="M1130 1170H120c-49.7 0-90-40.3-90-90V120c0-49.7 40.3-90 90-90h1010c49.7 0 90 40.3 90 90v960c0 49.7-40.3 90-90 90z" fill="#fff"/><path d="M1130 1140H120c-33.1 0-60-26.9-60-60V260h1130v820c0 33.1-26.9 60-60 60z" fill="#08963b"/><path d="M1190 260H60V120c0-33.1 26.9-60 60-60h1010c33.1 0 60 26.9 60 60z" fill="#e71f20"/><text xml:space="preserve" x="779.661" y="817.142" font-weight="bold" font-size="740" font-family="Source Han Sans" letter-spacing="0" fill="#fbf9f9"><tspan x="625" y="817.142" font-weight="normal" font-family="B型交通标志专用字体" text-anchor="middle">HIGHWAY_NUMBER</tspan></text><text x="625" y="1016.459" letter-spacing="6" fill="#fbf9f9" font-family="A型交通标志专用字体" font-size="133.333"><tspan text-anchor="middle">HIGHWAY_NAME</tspan></text><text x="625" y="190" letter-spacing="100" fill="#fbf9f9" font-family="A型交通标志专用字体" font-size="100"><tspan text-anchor="middle">国家高速</tspan></text></svg>''',
    
    "national_2digit_8char": '''NATIONAL_2DIGIT_8CHAR_TEMPLATE''',

    "national_4digit_4char": '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1700 1200" width="850" height="600" xmlns:v="https://vecta.io/nano"><path d="M1580 1200H120c-66.3 0-120-53.7-120-120V120C0 53.7 53.7 0 120 0h1460c66.3 0 120 53.7 120 120v960c0 66.3-53.7 120-120 120z" fill="#08963b"/><path d="M1580 1170H120c-49.7 0-90-40.3-90-90V120c0-49.7 40.3-90 90-90h1460c49.7 0 90 40.3 90 90v960c0 49.7-40.3 90-90 90z" fill="#fff"/><path d="M1580 1140H120c-33.1 0-60-26.9-60-60V260h1580v820c0 33.1-26.9 60-60 60z" fill="#08963b"/><path d="M1640 260H60V120c0-33.1 26.9-60 60-60h1460c33.1 0 60 26.9 60 60z" fill="#e71f20"/><text xml:space="preserve" x="1153.391" y="782.743" font-weight="bold" font-size="725.81" font-family="Source Han Sans" letter-spacing="0" fill="#fbf9f9"><tspan x="850" y="782.743" font-weight="normal" font-family="B型交通标志专用字体" text-anchor="middle">HIGHWAY_NUMBER_PART1<tspan font-size="477.335">HIGHWAY_NUMBER_PART2</tspan></tspan></text><text x="850" y="190" letter-spacing="200" fill="#fbf9f9" font-family="A型交通标志专用字体" font-size="100"><tspan text-anchor="middle">国家高速</tspan></text><text x="850" y="1020" letter-spacing="50" fill="#fbf9f9" font-family="A型交通标志专用字体" font-size="200"><tspan x="875" y="1020" text-anchor="middle">HIGHWAY_NAME</tspan></text></svg>''',

    "national_4digit_6char": '''NATIONAL_4DIGIT_6CHAR_TEMPLATE''',
    "national_4digit_7char": '''NATIONAL_4DIGIT_7CHAR_TEMPLATE''',
    "national_4digit_8char": '''NATIONAL_4DIGIT_8CHAR_TEMPLATE''',

    # 省级高速模板
    "provincial_1digit_noname": '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1000 1000" width="500" height="500" xmlns:v="https://vecta.io/nano"><path d="M880 1000H120C53.7 1000 0 946.3 0 880V120C0 53.7 53.7 0 120 0h760c66.3 0 120 53.7 120 120v760c0 66.3-53.7 120-120 120z" fill="#fff"/><path d="M880 970H120c-49.7 0-90-40.3-90-90V230h940v650c0 49.7-40.3 90-90 90z" fill="#08963b"/><path d="M970 230H30V120c0-49.7 40.3-90 90-90h760c49.7 0 90 40.3 90 90z" fill="#f6ec0a"/><text xml:space="preserve" x="651.451" y="812.755" font-weight="bold" font-size="724.642" font-family="Source Han Sans" letter-spacing="0" fill="#fbf9f9"><tspan x="500" y="812.755" font-weight="normal" font-family="B型交通标志专用字体" text-anchor="middle">HIGHWAY_NUMBER</tspan></text><text x="500" y="160" font-size="100" font-family="A型交通标志专用字体" letter-spacing="100" fill="#535353"><tspan text-anchor="middle">PROVINCE高速</tspan></text></svg>''',

    "provincial_2digit_noname": '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1250 1000" width="625" height="500" xmlns:v="https://vecta.io/nano"><path d="M1130 1000H120C53.7 1000 0 946.3 0 880V120C0 53.7 53.7 0 120 0h1010c66.3 0 120 53.7 120 120v760c0 66.3-53.7 120-120 120z" fill="#fff"/><path d="M1130 970H120c-49.7 0-90-40.3-90-90V230h1190v650c0 49.7-40.3 90-90 90z" fill="#08963b"/><path d="M1220 230H30V120c0-49.7 40.3-90 90-90h1010c49.7 0 90 40.3 90 90z" fill="#f6ec0a"/><text xml:space="preserve" x="776.451" y="812.755" font-weight="bold" font-size="724.642" font-family="Source Han Sans" letter-spacing="0" fill="#fbf9f9"><tspan x="625" y="812.755" font-weight="normal" font-family="B型交通标志专用字体" text-anchor="middle">HIGHWAY_NUMBER</tspan></text><text x="625" y="160" font-size="100" font-family="A型交通标志专用字体" letter-spacing="100" fill="#535353"><tspan text-anchor="middle">PROVINCE高速</tspan></text></svg>''',

    "provincial_4digit_noname": '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1700 1000" width="850" height="500" xmlns:v="https://vecta.io/nano"><path d="M1580 1000H120C53.7 1000 0 946.3 0 880V120C0 53.7 53.7 0 120 0h1460c66.3 0 120 53.7 120 120v760c0 66.3-53.7 120-120 120z" fill="#fff"/><path d="M1575 970H125c-49.7 0-90-40.3-90-90V230h1630v650c0 49.7-40.3 90-90 90z" fill="#08963b"/><path d="M1575 30H125c-49.7 0-90 40.3-90 90v110h1630V120c0-49.7-40.3-90-90-90z" fill="#f6ec0a"/><text xml:space="preserve" x="1001.207" y="812.766" font-weight="bold" font-size="723.475" font-family="Source Han Sans" letter-spacing="0" fill="#fbf9f9"><tspan x="850" y="812.766" font-weight="normal" font-family="B型交通标志专用字体" text-anchor="middle">HIGHWAY_NUMBER_PART1<tspan font-size="475.799">HIGHWAY_NUMBER_PART2</tspan></tspan></text><text x="850" y="160" font-size="100" font-family="A型交通标志专用字体" letter-spacing="191.199" fill="#535353"><tspan text-anchor="middle">PROVINCE高速</tspan></text></svg>''',


    "provincial_2digit_4char": '''<svg xmlns="http://www.w3.org/2000/svg" width="625" height="600" viewBox="0 0 1250 1200" xmlns:v="https://vecta.io/nano"><path d="M1130 1200H120c-66.3 0-120-53.7-120-120V120C0 53.7 53.7 0 120 0h1010c66.3 0 120 53.7 120 120v960c0 66.3-53.7 120-120 120z" fill="#08963b"/><path d="M1130 1170H120c-49.7 0-90-40.3-90-90V120c0-49.7 40.3-90 90-90h1010c49.7 0 90 40.3 90 90v960c0 49.7-40.3 90-90 90z" fill="#fff"/><path d="M1130 1140H120c-33.1 0-60-26.9-60-60V260h1130v820c0 33.1-26.9 60-60 60z" fill="#08963b"/><path d="M1190 260H60V120c0-33.1 26.9-60 60-60h1010c33.1 0 60 26.9 60 60z" fill="#f6ec0a"/><text xml:space="preserve" x="1111.235" y="782.755" font-size="724.642" letter-spacing="0" font-weight="bold" font-family="Source Han Sans" fill="#fbf9f9"><tspan x="625" y="782.755" font-weight="normal" font-family="B型交通标志专用字体" text-anchor="middle">HIGHWAY_NUMBER</tspan></text><text x="625" y="1020" letter-spacing="50" font-family="A型交通标志专用字体" fill="#fbf9f9" font-size="200"><tspan text-anchor="middle">HIGHWAY_NAME</tspan></text><text x="625" y="190" letter-spacing="100" font-family="A型交通标志专用字体" font-size="100" fill="#535353"><tspan text-anchor="middle">PROVINCE高速</tspan></text></svg>''',

    "provincial_2digit_5char": '''<svg xmlns="http://www.w3.org/2000/svg" width="625" height="600" viewBox="0 0 1250 1200" xmlns:v="https://vecta.io/nano"><path d="M1130 1200H120c-66.3 0-120-53.7-120-120V120C0 53.7 53.7 0 120 0h1010c66.3 0 120 53.7 120 120v960c0 66.3-53.7 120-120 120z" fill="#08963b"/><path d="M1130 1170H120c-49.7 0-90-40.3-90-90V120c0-49.7 40.3-90 90-90h1010c49.7 0 90 40.3 90 90v960c0 49.7-40.3 90-90 90z" fill="#fff"/><path d="M1130 1140H120c-33.1 0-60-26.9-60-60V260h1130v820c0 33.1-26.9 60-60 60z" fill="#08963b"/><path d="M1190 260H60V120c0-33.1 26.9-60 60-60h1010c33.1 0 60 26.9 60 60z" fill="#f6ec0a"/><text xml:space="preserve" x="1111.235" y="782.755" font-size="724.642" letter-spacing="0" font-weight="bold" font-family="Source Han Sans" fill="#fbf9f9"><tspan x="625" y="782.755" font-weight="normal" font-family="B型交通标志专用字体" text-anchor="middle">HIGHWAY_NUMBER</tspan></text><text x="625" y="1020" letter-spacing="0" font-family="A型交通标志专用字体" fill="#fbf9f9" font-size="200"><tspan letter-spacing="0" text-anchor="middle">HIGHWAY_NAME</tspan></text><text x="625" y="190" letter-spacing="100" font-family="A型交通标志专用字体" font-size="100" fill="#535353"><tspan text-anchor="middle">PROVINCE高速</tspan></text></svg>''',

    "provincial_2digit_6char": '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1250 1200" width="625" height="600" xmlns:v="https://vecta.io/nano"><path d="M1130 1200H120c-66.3 0-120-53.7-120-120V120C0 53.7 53.7 0 120 0h1010c66.3 0 120 53.7 120 120v960c0 66.3-53.7 120-120 120z" fill="#08963b"/><path d="M1130 1170H120c-49.7 0-90-40.3-90-90V120c0-49.7 40.3-90 90-90h1010c49.7 0 90 40.3 90 90v960c0 49.7-40.3 90-90 90z" fill="#fff"/><path d="M1130 1140H120c-33.1 0-60-26.9-60-60V260h1130v820c0 33.1-26.9 60-60 60z" fill="#08963b"/><path d="M1190 260H60V120c0-33.1 26.9-60 60-60h1010c33.1 0 60 26.9 60 60z" fill="#f6ec0a"/><text xml:space="preserve" x="775.965" y="800.167" font-size="722.313" letter-spacing="0" font-weight="bold" font-family="Source Han Sans" fill="#fbf9f9"><tspan x="625" y="800.167" font-weight="normal" font-family="B型交通标志专用字体" text-anchor="middle">HIGHWAY_NUMBER</tspan></text><text x="625" y="1014.679" letter-spacing="9.806" font-family="A型交通标志专用字体" fill="#fbf9f9" font-size="156.904"><tspan text-anchor="middle">HIGHWAY_NAME</tspan></text><text x="625" y="190" letter-spacing="100" font-family="A型交通标志专用字体" font-size="100" fill="#535353"><tspan text-anchor="middle">PROVINCE高速</tspan></text></svg>''',


    "provincial_2digit_7char": '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1250 1200" width="625" height="600" xmlns:v="https://vecta.io/nano"><path d="M1130 1200H120c-66.3 0-120-53.7-120-120V120C0 53.7 53.7 0 120 0h1010c66.3 0 120 53.7 120 120v960c0 66.3-53.7 120-120 120z" fill="#08963b"/><path d="M1130 1170H120c-49.7 0-90-40.3-90-90V120c0-49.7 40.3-90 90-90h1010c49.7 0 90 40.3 90 90v960c0 49.7-40.3 90-90 90z" fill="#fff"/><path d="M1130 1140H120c-33.1 0-60-26.9-60-60V260h1130v820c0 33.1-26.9 60-60 60z" fill="#08963b"/><path d="M1190 260H60V120c0-33.1 26.9-60 60-60h1010c33.1 0 60 26.9 60 60z" fill="#f6ec0a"/><text xml:space="preserve" x="934.322" y="817.142" font-weight="bold" font-size="740" font-family="Source Han Sans" letter-spacing="0" fill="#fbf9f9"><tspan x="625" y="817.142" font-weight="normal" font-family="B型交通标志专用字体" text-anchor="middle">HIGHWAY_NUMBER</tspan></text><text x="625" y="1016.459" letter-spacing="6" fill="#fbf9f9" font-family="A型交通标志专用字体" font-size="133.333"><tspan text-anchor="middle">HIGHWAY_NAME</tspan></text><text x="625" y="190" letter-spacing="100" fill="#535353" font-family="A型交通标志专用字体" font-size="100"><tspan text-anchor="middle">PROVINCE高速</tspan></text></svg>''',

    "provincial_2digit_8char": '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1250 1200" width="625" height="600" xmlns:v="https://vecta.io/nano"><path d="M1130 1200H120c-66.3 0-120-53.7-120-120V120C0 53.7 53.7 0 120 0h1010c66.3 0 120 53.7 120 120v960c0 66.3-53.7 120-120 120z" fill="#08963b"/><path d="M1130 1170H120c-49.7 0-90-40.3-90-90V120c0-49.7 40.3-90 90-90h1010c49.7 0 90 40.3 90 90v960c0 49.7-40.3 90-90 90z" fill="#fff"/><path d="M1130 1140H120c-33.1 0-60-26.9-60-60V260h1130v820c0 33.1-26.9 60-60 60z" fill="#08963b"/><path d="M1190 260H60V120c0-33.1 26.9-60 60-60h1010c33.1 0 60 26.9 60 60z" fill="#f6ec0a"/><text xml:space="preserve" x="934.322" y="824.919" font-weight="bold" font-size="740" font-family="Source Han Sans" letter-spacing="0" fill="#fbf9f9"><tspan x="625" y="824.919" font-weight="normal" font-family="B型交通标志专用字体" text-anchor="middle">HIGHWAY_NUMBER</tspan></text><text x="625" y="1008.664" letter-spacing="6" fill="#fbf9f9" font-family="A型交通标志专用字体" font-size="115.912"><tspan text-anchor="middle">HIGHWAY_NAME</tspan></text><text x="625" y="190" letter-spacing="100" fill="#535353" font-family="A型交通标志专用字体" font-size="100"><tspan text-anchor="middle">PROVINCE高速</tspan></text></svg>''',

    "provincial_4digit_4char": '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1700 1200" width="850" height="600" xmlns:v="https://vecta.io/nano"><path d="M1580 1200H120c-66.3 0-120-53.7-120-120V120C0 53.7 53.7 0 120 0h1460c66.3 0 120 53.7 120 120v960c0 66.3-53.7 120-120 120z" fill="#08963b"/><path d="M1580 1170H120c-49.7 0-90-40.3-90-90V120c0-49.7 40.3-90 90-90h1460c49.7 0 90 40.3 90 90v960c0 49.7-40.3 90-90 90z" fill="#fff"/><path d="M1580 1140H120c-33.1 0-60-26.9-60-60V260h1580v820c0 33.1-26.9 60-60 60z" fill="#08963b"/><path d="M1640 260H60V120c0-33.1 26.9-60 60-60h1460c33.1 0 60 26.9 60 60z" fill="#f6ec0a"/><text xml:space="preserve" x="1000.965" y="782.778" font-size="722.315" letter-spacing="0" font-weight="bold" font-family="Source Han Sans" fill="#fbf9f9"><tspan x="850" y="782.778" font-weight="normal" font-family="B型交通标志专用字体" text-anchor="middle">HIGHWAY_NUMBER_PART1<tspan font-size="475.036">HIGHWAY_NUMBER_PART2</tspan></tspan></text><text x="850" y="190" letter-spacing="191.199" font-family="A型交通标志专用字体" font-size="100" fill="#535353"><tspan text-anchor="middle">PROVINCE高速</tspan></text><text x="850" y="1020" letter-spacing="52.192" font-family="A型交通标志专用字体" fill="#fbf9f9" font-size="200"><tspan text-anchor="middle">HIGHWAY_NAME</tspan></text></svg>''',

    "provincial_4digit_6char": '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1700 1200" width="850" height="600" xmlns:v="https://vecta.io/nano"><path d="M1580 1200H120c-66.3 0-120-53.7-120-120V120C0 53.7 53.7 0 120 0h1460c66.3 0 120 53.7 120 120v960c0 66.3-53.7 120-120 120z" fill="#08963b"/><path d="M1580 1170H120c-49.7 0-90-40.3-90-90V120c0-49.7 40.3-90 90-90h1460c49.7 0 90 40.3 90 90v960c0 49.7-40.3 90-90 90z" fill="#fff"/><path d="M1580 1140H120c-33.1 0-60-26.9-60-60V260h1580v820c0 33.1-26.9 60-60 60z" fill="#08963b"/><path d="M1640 260H60V120c0-33.1 26.9-60 60-60h1460c33.1 0 60 26.9 60 60z" fill="#f6ec0a"/><text xml:space="preserve" x="1000.965" y="782.778" font-size="722.315" letter-spacing="0" font-weight="bold" font-family="Source Han Sans" fill="#fbf9f9"><tspan x="850" y="782.778" font-weight="normal" font-family="B型交通标志专用字体" text-anchor="middle">HIGHWAY_NUMBER_PART1<tspan font-size="475.036">HIGHWAY_NUMBER_PART2</tspan></tspan></text><text x="850" y="190" letter-spacing="191.199" font-family="A型交通标志专用字体" font-size="100" fill="#535353"><tspan text-anchor="middle">PROVINCE高速</tspan></text><text x="850" y="1020" letter-spacing="0" font-family="A型交通标志专用字体" fill="#fbf9f9" font-size="200"><tspan letter-spacing="0" text-anchor="middle">HIGHWAY_NAME</tspan></text></svg>''',
    "provincial_4digit_8char": '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1700 1200" width="850" height="600" xmlns:v="https://vecta.io/nano"><path d="M1580 1200H120c-66.3 0-120-53.7-120-120V120C0 53.7 53.7 0 120 0h1460c66.3 0 120 53.7 120 120v960c0 66.3-53.7 120-120 120z" fill="#08963b"/><path d="M1580 1170H120c-49.7 0-90-40.3-90-90V120c0-49.7 40.3-90 90-90h1460c49.7 0 90 40.3 90 90v960c0 49.7-40.3 90-90 90z" fill="#fff"/><path d="M1580 1140H120c-33.1 0-60-26.9-60-60V260h1580v820c0 33.1-26.9 60-60 60z" fill="#08963b"/><path d="M1640 260H60V120c0-33.1 26.9-60 60-60h1460c33.1 0 60 26.9 60 60z" fill="#f6ec0a"/><text xml:space="preserve" x="1000.965" y="782.778" font-size="722.315" letter-spacing="0" font-weight="bold" font-family="Source Han Sans" fill="#fbf9f9"><tspan x="850" y="782.778" font-weight="normal" font-family="B型交通标志专用字体" text-anchor="middle">HIGHWAY_NUMBER_PART1<tspan font-size="475.036">HIGHWAY_NUMBER_PART2</tspan></tspan></text><text x="850" y="190" letter-spacing="191.199" font-family="A型交通标志专用字体" font-size="100" fill="#535353"><tspan text-anchor="middle">PROVINCE高速</tspan></text><text x="850" y="1020" letter-spacing="0" font-family="A型交通标志专用字体" fill="#fbf9f9" font-size="173.333"><tspan letter-spacing="0" text-anchor="middle">HIGHWAY_NAME</tspan></text></svg>''',
}

TEMPLATES["national_1digit_4char"] = TEMPLATES["national_2digit_4char"]
TEMPLATES["national_1digit_5char"] = TEMPLATES["national_2digit_5char"]
TEMPLATES["national_1digit_6char"] = TEMPLATES["national_2digit_6char"]
TEMPLATES["national_1digit_7char"] = TEMPLATES["national_2digit_7char"]
TEMPLATES["national_1digit_8char"] = TEMPLATES["national_2digit_8char"]
TEMPLATES["national_4digit_5char"] = TEMPLATES["national_4digit_4char"]
TEMPLATES["provincial_1digit_4char"] = TEMPLATES["provincial_2digit_4char"]
TEMPLATES["provincial_1digit_5char"] = TEMPLATES["provincial_2digit_5char"]
TEMPLATES["provincial_1digit_6char"] = TEMPLATES["provincial_2digit_6char"]
TEMPLATES["provincial_1digit_7char"] = TEMPLATES["provincial_2digit_7char"]
TEMPLATES["provincial_1digit_8char"] = TEMPLATES["provincial_2digit_8char"]
TEMPLATES["provincial_4digit_5char"] = TEMPLATES["provincial_4digit_4char"]
TEMPLATES["provincial_4digit_7char"] = TEMPLATES["provincial_4digit_6char"]


def get_template_key(province, highway_number, highway_name):
    """根据条件确定使用的模板key"""
    # 判断是国家高速还是省级高速
    is_national = (province == "国家")

    # 判断数字位数（使用总字符数-1）
    if pd.isna(highway_number) or highway_number == "":
        digit_count = 0
    else:
        # 总字符数减1作为数字位数
        digit_count = len(str(highway_number)) - 1

    # 确保digit_count不为负数
    digit_count = max(0, digit_count)

    # 判断名称长度
    if pd.isna(highway_name) or highway_name == "":
        name_length = "noname"
    else:
        name_length = f"{len(str(highway_name))}char"

    # 构建模板key
    highway_type = "national" if is_national else "provincial"
    template_key = f"{highway_type}_{digit_count}digit_{name_length}"

    return template_key


def create_highway_sign(province, highway_number, highway_name, save_path):
    """创建高速标志SVG文件"""
    template_key = get_template_key(province, highway_number, highway_name)

    if template_key not in TEMPLATES or TEMPLATES[template_key].startswith(template_key.split('_')[0].upper()):
        raise ValueError(f"不支持此组合：省份='{province}', 编号='{highway_number}', 名称='{highway_name}'")

    svg_content = TEMPLATES[template_key]

    # 替换高速编号
    if "HIGHWAY_NUMBER_PART1" in svg_content and "HIGHWAY_NUMBER_PART2" in svg_content:
        # 4位数字的情况，需要分割
        if len(highway_number) >= 3:
            part1 = highway_number[:3]  # 前3位，如"S27"
            part2 = highway_number[3:]  # 剩余部分，如"14"
            svg_content = svg_content.replace('HIGHWAY_NUMBER_PART1', part1)
            svg_content = svg_content.replace('HIGHWAY_NUMBER_PART2', part2)
        else:
            # 如果长度不足3位，使用完整编号
            svg_content = svg_content.replace('HIGHWAY_NUMBER_PART1', highway_number)
            svg_content = svg_content.replace('HIGHWAY_NUMBER_PART2', '')
    else:
        # 普通编号替换
        svg_content = svg_content.replace('HIGHWAY_NUMBER', highway_number)

    # 替换省份
    if province != "国家":
        svg_content = svg_content.replace('PROVINCE', province)

    # 替换高速名称
    if highway_name and str(highway_name) != "nan":
        name_str = str(highway_name)
        svg_content = svg_content.replace('HIGHWAY_NAME', name_str)

    # 保存SVG文件
    with open(save_path, 'w', encoding='utf-8') as f:
        f.write(svg_content)


def main():
    # 创建保存目录
    os.makedirs(SAVE_DIR, exist_ok=True)

    # 存储成功生成的文件路径
    success_files = []

    try:
        # 检查Excel文件是否存在
        if not os.path.exists(EXCEL_PATH):
            raise FileNotFoundError(f"找不到Excel文件 '{EXCEL_PATH}'")

        # 读取Excel文件（无表头）
        df = pd.read_excel(EXCEL_PATH, header=None)

        # 检查数据是否为空
        if df.empty:
            print("Excel文件为空，没有数据需要处理")
            return

        # 确保DataFrame至少有3列
        if len(df.columns) < 3:
            # 如果列数不足，添加空列
            for i in range(len(df.columns), 3):
                df[i] = None

        # 统计处理情况
        total_count = len(df)
        success_count = 0
        skip_count = 0
        failed_rows = []

        print(f"开始处理Excel文件，共{total_count}行数据...")

        for index, row in df.iterrows():
            # 安全地获取三列数据，使用try-except处理可能的列不存在问题
            try:
                province = str(row[0]).strip() if pd.notna(row[0]) else ""  # 第一列：省份
            except KeyError:
                province = ""

            try:
                highway_number = str(row[1]).strip() if pd.notna(row[1]) else ""  # 第二列：高速编号
            except KeyError:
                highway_number = ""

            try:
                highway_name = row[2]  # 第三列：高速名称（可能为NaN）
            except KeyError:
                highway_name = None

            # 基本校验
            if not province:
                print(f"第{index + 1}行跳过：省份为空")
                failed_rows.append((index + 1, province, highway_number, highway_name, "省份为空"))
                skip_count += 1
                continue

            if not highway_number:
                print(f"第{index + 1}行跳过：高速编号为空")
                failed_rows.append((index + 1, province, highway_number, highway_name, "高速编号为空"))
                skip_count += 1
                continue

            # 生成文件名
            if NAMING_STYLE == 1:
                # 中文命名方式
                if pd.isna(highway_name) or not highway_name:
                    filename = f"{province}高速_{highway_number}_无名称.svg"
                else:
                    filename = f"{province}高速_{highway_number}_{highway_name}.svg"
            else:
                # 英文命名方式
                province_eng = PROVINCE_ENGLISH.get(province, "Unknown")
                if province == "国家":
                    # 国家高速
                    if pd.isna(highway_name) or not highway_name:
                        filename = f"China Expwy {highway_number} sign no name.svg"
                    else:
                        filename = f"China Expwy {highway_number} sign with name.svg"
                else:
                    # 省级高速
                    if pd.isna(highway_name) or not highway_name:
                        filename = f"{province_eng} Expwy {highway_number} sign no name.svg"
                    else:
                        filename = f"{province_eng} Expwy {highway_number} sign with name.svg"

            save_path = os.path.join(SAVE_DIR, filename)

            try:
                # 创建SVG文件
                create_highway_sign(province, highway_number, highway_name, save_path)
                template_key = get_template_key(province, highway_number, highway_name)
                print(f"第{index + 1}行成功：{filename}（模板：{template_key}）")
                success_count += 1
                success_files.append(save_path)  # 记录成功生成的文件路径

            except ValueError as e:
                print(f"第{index + 1}行跳过：{str(e)}")
                failed_rows.append((index + 1, province, highway_number, highway_name, str(e)))
                skip_count += 1
            except Exception as e:
                print(f"第{index + 1}行错误：创建文件失败 - {str(e)}")
                failed_rows.append((index + 1, province, highway_number, highway_name, f"创建失败：{str(e)}"))
                skip_count += 1

        # 对成功生成的文件进行文字转曲处理
        if success_files:
            convert_text_to_path(success_files, INKSCAPE_PATH)
        else:
            print("没有成功生成的SVG文件，跳过文字转曲处理")

        # 输出统计信息
        print(f"\n处理完成！")
        print(f"总计：{total_count}行")
        print(f"成功：{success_count}个SVG文件")
        print(f"跳过：{skip_count}行")
        print(f"保存目录：{SAVE_DIR}")

        # 输出未能处理的行
        if failed_rows:
            print(f"\n未能处理的行：")
            for row_num, prov, num, name, reason in failed_rows:
                print(f"  第{row_num}行：省份='{prov}', 编号='{num}', 名称='{name}' - 原因：{reason}")

    except FileNotFoundError as e:
        print(f"错误：{str(e)}")
    except pd.errors.EmptyDataError:
        print("错误：Excel文件为空")
    except pd.errors.ParserError as e:
        print(f"错误：Excel文件解析失败 - {str(e)}")
    except PermissionError:
        print(f"错误：没有权限读取Excel文件 '{EXCEL_PATH}'")
    except Exception as e:
        print(f"读取Excel文件时发生错误：{str(e)}")
        print(f"错误类型：{type(e).__name__}")


def convert_text_to_path(success_files, inkscape_path):
    """将SVG文件中的文字转换为曲线路径"""
    import subprocess

    if not success_files:
        print("没有需要处理的SVG文件")
        return

    print(f"\n开始文字转曲处理，共{len(success_files)}个文件...")

    processed_count = 0
    failed_files = []

    for svg_file in success_files:
        try:
            filename = os.path.basename(svg_file)

            # 构建Inkscape命令
            actions = [
                f"export-filename:{svg_file}",  # 覆盖原文件
                "export-do"
            ]

            command = [
                inkscape_path,
                f"--actions={';'.join(actions)}",
                "--export-text-to-path",
                svg_file
            ]

            # 执行命令
            result = subprocess.run(command, capture_output=True, text=True, timeout=60)

            if result.returncode == 0:
                print(f"  已处理: {filename}")
                processed_count += 1
            else:
                print(f"  处理失败: {filename}")
                failed_files.append((filename, result.stderr))

        except subprocess.TimeoutExpired:
            print(f"  处理超时: {filename}")
            failed_files.append((filename, "处理超时"))
        except Exception as e:
            print(f"  处理异常: {filename} - {str(e)}")
            failed_files.append((filename, str(e)))

    # 输出处理结果
    print(f"\n文字转曲处理完成:")
    print(f"  成功处理: {processed_count} 个文件")
    if failed_files:
        print(f"  处理失败: {len(failed_files)} 个文件")
        for filename, error in failed_files:
            print(f"    {filename}: {error}")


if __name__ == "__main__":
    main()
