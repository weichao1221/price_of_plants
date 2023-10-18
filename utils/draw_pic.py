from data import draw, get_name_list
import time

start_time = time.time()
with open('log.txt', 'w') as log_file:
    start_name_list = get_name_list()
    for start_name in start_name_list:
        start_time = time.time()
        print(f"{start_name}开始绘制")
        try:
            draw(start_name)
            log_file.write(f'{start_name}已完成绘制，{time.strftime("%Y-%m-%d %H-%M-%S")}\n')
        except Exception as e:
            error_message = f'{start_name}绘制失败，原因：{e}\n'
            print(error_message)
            log_file.write(error_message)
            continue
        # time.sleep(2)
        end_time = time.time()
        print(f'{start_name}绘制完成，耗时{round(end_time - start_time, 2)}秒')

end_time = time.time()
print(f'所有内容绘制完成，耗时{round(end_time - start_time, 2)}秒')
input('按任意键退出...')