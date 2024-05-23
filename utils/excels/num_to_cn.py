def number_to_chinese(num):
    chinese_number_map = {
        0: '零',
        1: '壹',
        2: '贰',
        3: '叁',
        4: '肆',
        5: '伍',
        6: '陆',
        7: '柒',
        8: '捌',
        9: '玖',
    }
    chinese_unit_map = {
        10: '拾',
        100: '佰',
        1000: '仟',
        10000: '万',
        100000000: '亿',
        1000000000000: '兆',
        10000000000000000: '京',
        100000000000000000000: '垓',
        1000000000000000000000000: '秭',
        10000000000000000000000000000: '穰',
        100000000000000000000000000000000: '沟',
    }

    def convert_chunk(chunk):
        result = ''
        length = len(chunk)
        for i, digit in enumerate(chunk):
            digit = int(digit)
            if digit != 0:
                result += chinese_number_map[digit]
                if i < length - 1:
                    unit = 10 ** (length - i - 1)
                    if unit in chinese_unit_map:
                        result += chinese_unit_map[unit]
                    elif unit % 1000 == 0:
                        result += '仟'
                    elif unit % 100 == 0:
                        result += '佰'
                    elif unit % 10 == 0:
                        result += '拾'
            else:
                # 处理单独的零
                if i < length - 1 and chunk[i + 1] != '0':
                    result += chinese_number_map[digit]

        return result

    def convert_integer_part(num_str):
        if num_str == '0':
            return chinese_number_map[0]

        length = len(num_str)

        # 处理整数部分
        result = ''
        if length <= 4:
            result = convert_chunk(num_str)
        elif length <= 8:
            result = convert_chunk(num_str[:length - 4]) + '万' + convert_chunk(num_str[length - 4:])
        elif length <= 12:
            result = convert_chunk(num_str[:length - 8]) + '亿' + convert_chunk(num_str[length - 8:length - 4]) + '万' + convert_chunk(num_str[length - 4:])
        elif length <= 16:
            result = convert_chunk(num_str[:length - 12]) + '兆' + convert_chunk(num_str[length - 12:length - 8]) + '亿' + convert_chunk(num_str[length - 8:length - 4]) + '万' + convert_chunk(num_str[length - 4:])
        elif length <= 20:
            result = convert_chunk(num_str[:length - 16]) + '京' + convert_chunk(num_str[length - 16:length - 12]) + '兆' + convert_chunk(num_str[length - 12:length - 8]) + '亿' + convert_chunk(num_str[length - 8:length - 4]) + '万' + convert_chunk(num_str[length - 4:])
        elif length <= 24:
            result = convert_chunk(num_str[:length - 20]) + '垓' + convert_chunk(num_str[length - 20:length - 16]) + '京' + convert_chunk(num_str[length - 16:length - 12]) + '兆' + convert_chunk(num_str[length - 12:length - 8]) + '亿' + convert_chunk(num_str[length - 8:length - 4]) + '万' + convert_chunk(num_str[length - 4:])
        elif length <= 28:
            result = convert_chunk(num_str[:length - 24]) + '秭' + convert_chunk(num_str[length - 24:length - 20]) + '垓' + convert_chunk(num_str[length - 20:length - 16]) + '京' + convert_chunk(num_str[length - 16:length - 12]) + '兆' + convert_chunk(num_str[length - 12:length - 8]) + '亿' + convert_chunk(num_str[length - 8:length - 4]) + '万' + convert_chunk(num_str[length - 4:])
        elif length <= 32:
            result = convert_chunk(num_str[:length - 28]) + '穰' + convert_chunk(num_str[length - 28:length - 24]) + '秭' + convert_chunk(num_str[length - 24:length - 20]) + '垓' + convert_chunk(num_str[length - 20:length - 16]) + '京' + convert_chunk(num_str[length - 16:length - 12]) + '兆' + convert_chunk(num_str[length - 12:length - 8]) + '亿' + convert_chunk(num_str[length - 8:length - 4]) + '万' + convert_chunk(num_str[length - 4:])
        elif length <= 36:
            result = convert_chunk(num_str[:length - 32]) + '沟' + convert_chunk(num_str[length - 32:length - 28]) + '穰' + convert_chunk(num_str[length - 28:length - 24]) + '秭' + convert_chunk(num_str[length - 24:length - 20]) + '垓' + convert_chunk(num_str[length - 20:length - 16]) + '京' + convert_chunk(num_str[length - 16:length - 12]) + '兆' + convert_chunk(num_str[length - 12:length - 8]) + '亿' + convert_chunk(num_str[length - 8:length - 4]) + '万' + convert_chunk(num_str[length - 4:])
        else:
            result = '超出范围'
        return result

    def convert_decimal_part(decimal_str):
        result = ''
        if not decimal_str or decimal_str == "0" or decimal_str == "00":
            result = "整"
        else:
            if int(decimal_part[0]) == 0:
                result += "零"
            else:
                result += chinese_number_map[int(decimal_part[0])] + "角"
            if len(decimal_part) > 1 and decimal_part[1] != "0":
                result += chinese_number_map[int(decimal_part[1])] + "分"
            else:
                result += "整"
        return result

    num_str = str(num)
    if '.' in num_str:
        integer_part, decimal_part = num_str.split('.') # 分割整数和小数部分
        print(f"整数部分: {integer_part}", f"小数部分: {decimal_part}")
        integer_chinese = convert_integer_part(integer_part)    # 转换整数部分
        decimal_chinese = convert_decimal_part(decimal_part)    # 转换小数部分

        if decimal_chinese:
            result = integer_chinese + '元' + decimal_chinese    # 拼接整数和小数部分
        else:
            result = integer_chinese + '元整'   # 拼接整数和小数部分
    else:
        result = convert_integer_part(num_str) + "元整"

    return result

# 测试示例
if __name__ == "__main__":
    num = 12001
    num1 = 999999999999.00
    num2 = 12001.00
    num3 = 1120012001.10

    print(f"{num},{number_to_chinese(num)}")
    print(f"{num1},{number_to_chinese(num1)}")
    print(f"{num2},{number_to_chinese(num2)}")
    print(f"{num3},{number_to_chinese(num3)}")
