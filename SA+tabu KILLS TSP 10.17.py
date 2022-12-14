'''
2022.10.17版本
'''



import math
import random
import matplotlib.pyplot as plt
import xlrd
import xlwt

T_begin = 60000      #初始温度
T_end = 5      #结束温度
alpha = 0.99        #降温系数
L = 3000        #迭代次数
buildings = 19  #楼栋数量
Demand_matrix_root = "D:\demo\yiwenchao finnal course\Demand matrix.xlsx"       #修改成Demand matrix.xlsx在自己电脑的路径
Distance_matrix_root = "D:\demo\yiwenchao finnal course\Distance matrix.xlsx"   #同上

#表格数据读取函数 输入：需求文件路径，距离文件路径
def read_excel_data(demand_root,distance_root):
    Demand_matrix_table = []                                                    #创建需求矩阵表
    Distance_matrix_table = []                                                  #创建距离矩阵表
    data_demand = xlrd.open_workbook(demand_root)                               #打开需求文件
    sheet_demand = data_demand.sheet_by_index(0)                                #打开需求文件的第一个表
    data_distance = xlrd.open_workbook(distance_root)                           #打开距离文件
    sheet_distance = data_distance.sheet_by_index(0)                            #打开距离文件的第一个表
    for i in range(sheet_demand.nrows):                                         #按行读取需求数据
        Demand_matrix_table.append(sheet_demand.row_values(i))                  #保存行需求数据于需求矩阵表中
    for i in range(sheet_distance.nrows):                                       #按行读取距离数据
        Distance_matrix_table.append(sheet_distance.row_values(i))              #保存行距离数据于距离矩阵表中
    return Demand_matrix_table,Distance_matrix_table                            #输出需求矩阵表和距离矩阵表



#两个城市的距离 输入：两点、距离矩阵
def dist(a,b,Distance_matrix):
    return Distance_matrix[a][b]

#路程总长 输入：路线
def totaldistance(travel_route,Distance_matrix):
    value = 0
    for j in range(len(travel_route)-1):                                        #路径是1~19，下标是0~18，j = 0~17
        value = value + dist(travel_route[j],travel_route[j+1],Distance_matrix)
    value = value + dist(travel_route[len(travel_route)-1], travel_route[0],Distance_matrix)
    return value

#初始化一个解 [1,2,3,..,19] 19栋楼
def init_route():
    route = []
    for i in range(1,buildings+1):                                              #i = 1 ~ 19
        route.append(i)
    return route

#产生新解 输入：旧路径，禁忌列表
def creat_newRoute(old_route,tabu_list):
    new_route = []                                                              #新路线
    for i in range(len(old_route)):                                             #len(old_route) = 19,i = 0~18
        new_route.append(old_route[i])                                          #把旧路线复制给新路线
    while True:
        while True:
            change_a = random.randint(0,len(old_route)-1)                           #change_a = 0~18,random.randint(a,b)的取值范围是[a,b]
            change_b = random.randint(0,len(old_route)-1)                           #change_b = 0~18
            if change_a != change_b:                                            #保证a,b两点不同
                break
        new_route[change_a], new_route[change_b] = new_route[change_b], new_route[change_a]  # 交换两点
        if new_route not in tabu_list:                                          #判断新路线是否在禁忌列表中
            break
    tabu_list.append(new_route)                                                   #更新禁忌列表
    # print("禁忌列表：",tabu_list)
    return new_route,tabu_list                                                              #输出新路线和禁忌列表

#某路线的需求满足量计算函数 输入：车辆行驶速度、行驶路径、距离矩阵、需求矩阵
def totaldemand(v,travel_route,Distance_matrix,Demand_matrix):
    current_time = 0                                                            #当前时间为0
    sum_demand = Demand_matrix[travel_route[0]][current_time+1]                 #初始需求为起点在第0时刻的需求
    for i in range(len(travel_route)-1):                                        #i = 0~17
        frontpoint = int(travel_route[i])                                       #记录相邻两点的前点
        latterpoint = int(travel_route[i+1])                                    #记录相邻两点的后点
        distance_between = Distance_matrix[frontpoint][latterpoint]             #记录两点之间的距离
        time_cost = int(distance_between/v)                                     #计算从前点到后点的所需的时间
        latterdemand = Demand_matrix[latterpoint][current_time+time_cost+1]     #计算车辆到达时后点的需求
        sum_demand = sum_demand + latterdemand                                  #计算总需求满足量
        current_time += time_cost                                               #更新当前时间
    return sum_demand                                                           #输出总需求满足量

#计算适应度函数(合理吗？？？？？)
def fitness(route_distance, sum_demand):
    di = (route_distance - 1212)/(5306 - 1212)
    de = (sum_demand - 94)/(255 - 94)
    f = (0.3 * di - 0.7 * de)*1000
    return f

#模拟退火算法
def SA(T_begin,T_end,alpha,L,Distance_matrix,Demand_matrix):                    #T_begin初始温度;T_end终止温度;alpha降温系数;L迭代次数
    tabu_list = []                                                              #空禁忌列表
    old_route = init_route()                                                    #产生初始路径route0赋值给old_route
    T = T_begin                                                                 #当前温度等于初始温度
    cnt = 0                                                                     #计数器 用于输出报文时的序号 “第 cnt 次 ...”
    Shortestlength_list = []                                                    #记录每种温度情况下得到的最短路径距离
    Fitness_list = []                                                           #记录每种温度情况下得到的最小适应度
    best_fitness_list = []                                                      #记录历史最优适应度的列表
    Route_list = []                                                             #记录每种温度情况下得到的最小适应度对应的路线
    while T > T_end:                                                            #如果当前温度大于终止温度
        for i in range(L):                                                      #循环L次
            new_route,tabu_list = creat_newRoute(old_route,tabu_list)           #扰动产生新路径
            if len(tabu_list) > 0:                                              #如果禁忌列表长度≥4，那么清空禁忌列表
                tabu_list.pop(0)
            old_dist = totaldistance(old_route,Distance_matrix)                 #计算旧路径的长度
            new_dist = totaldistance(new_route,Distance_matrix)                 #计算新路径的长度
            old_fitness = fitness(old_dist,totaldemand(10,old_route,Distance_matrix,Demand_matrix))
            new_fitness = fitness(new_dist,totaldemand(10,new_route,Distance_matrix,Demand_matrix))
            delta_fitness = new_fitness - old_fitness
            if delta_fitness >= 0:                                              #如果差值≥0，说明新路线比旧路线长，按一定概率接受新路线
                rand = random.uniform(0,1)
                if rand < math.exp(-delta_fitness/T):
                    old_route = new_route
            else:                                                               #否则，必然接受新路线
                old_route = new_route
        T = T * alpha                                                           #降温
        cnt += 1                                                                #计数器
        Shortestlength = totaldistance(old_route,Distance_matrix)               #当前路径长度（更新后的旧路径）
        Shortestlength_list.append(Shortestlength)                              #把当前路径长度加入列表
        Fitness_list.append(fitness(Shortestlength,totaldemand(10,old_route,Distance_matrix,Demand_matrix)))
                                                                                #把当前适应度加入列表
        Route_list.append(old_route)                                            #把当前适应度对应的路线加入列表
        best_fitness_list.append(min(Fitness_list))  # 记录当前历史最优适应度，用于作图
        print(cnt,"次降温，温度为：",T," 路程长度为：", Shortestlength,"适应度：",fitness(Shortestlength,totaldemand(10,old_route,Distance_matrix,Demand_matrix)))
                                                                                #报文
    optimal_route = old_route                                                   #当温度降到终止温度，把此时的旧路线作为最优路线
    optimal_distance = totaldistance(optimal_route,Distance_matrix)             #计算最优路线的路程长度
    optimal_demand = totaldemand(10,old_route,Distance_matrix,Demand_matrix)
    optimal_fitness = fitness(optimal_distance,optimal_demand)
    print("本次独立运行 最优路径长度：",optimal_distance)                            #打印最优路径长度
    print("本次独立运行 最优路线：",optimal_route)                                  #打印最优路线
    print("本次独立运行 满足需求量：",optimal_demand)                               #打印需求满足量
    print("本次独立运行 适应度：", optimal_fitness)                                   # 打印适应度
    # plt.plot(Fitness_list)                                                      #绘制适应度曲线，如果不想波动曲线与折线重合，就把这条注释掉
    # plt.savefig('SA+tabu fitness (5 times)' + str(T_begin) + ' ' + str(T_end) + ' ' + str(alpha) + ' ' + str(L) + '.png')#保存到项目目录下命名为XXX
    best_fitness = min(Fitness_list)                #导出历史最优适应度
    plt.plot(best_fitness_list)                                                 #绘制历史最优适应度曲线
    plt.savefig('SA+tabu best fitness (5 times)' + str(T_begin) + ' ' + str(T_end) + ' ' + str(alpha) + ' ' + str(L) + '.png')
    index = Fitness_list.index(best_fitness)        #记录最优适应度的下标
    best_route = Route_list[index]                  #导出对应下表的路线
    return optimal_distance,optimal_route,optimal_demand,optimal_fitness,best_fitness,best_route                 #输出最优路程长度和最优路径

def main():
    Demand_matrix, Distance_matrix = read_excel_data(Demand_matrix_root, Distance_matrix_root)  #基础数据导入
    independent_runs = 5                                                       #独立运行次数
    Several_independent_runs_distance_list = []                                 #若干次独立运行得到的最优路程长度表
    Several_independent_runs_route_list = []                                    #若干次独立运行得到的最优路径表
    Several_independent_runs_fitness_list = []                                  #若干次独立运行得到的最优适应度函数表
    work_book = xlwt.Workbook(encoding='utf-8')                                 #创建工作表
    sheet = work_book.add_sheet('1')                                            #添加sheet表，命名为XXX
    sheet.write(0,0,"独立运行次数")                                               #第0行第0列，赋值XXX
    sheet.write(0,1,"最优路程长度")                                               #第0行第1列，赋值XXX
    sheet.write(0,2,"最优路径")                                                  #第0行第2列，赋值XXX
    sheet.write(0,3,"需求满足量")
    sheet.write(0,4,"适应度")
    sheet.write(0,5,"历史最优适应度")
    sheet.write(0,6,"历史最优路线")
    for i in range(independent_runs):                                           #独立运行 independent_runs 次
        optimal_distance,optimal_route,optimal_demand,optimal_fitness,best_fitness,best_route = SA(T_begin,T_end,alpha,L,Distance_matrix,Demand_matrix)#运行SA()，输出当次运行的最优长度和最优路线
        Several_independent_runs_distance_list.append(optimal_distance)         #记录当此运行的最优长度于列表中
        Several_independent_runs_route_list.append(optimal_route)               #记录当次运行的最优路径于列表中
        Several_independent_runs_fitness_list.append(optimal_fitness)           #记录当次运行的最优适应度函数于列表中
        sheet.write(i+1,0,str(i+1))                                             #第 i+1 行第 0 列输入当次迭代的序号
        sheet.write(i+1,1,optimal_distance)                                     #第 i+1 行第 1 列输入当次迭代的最优长度
        sheet.write(i+1,2,str(optimal_route))                                   #第 i+1 行第 2 列输入当次迭代的最优路线
        sheet.write(i+1,3,str(optimal_demand))
        sheet.write(i+1,4,str(optimal_fitness))
        sheet.write(i+1,5,str(best_fitness))
        sheet.write(i+1,6,str(best_route))
    work_book.save('SA + Tabu ('+str(independent_runs)+' times).xls')                                   #保存工作表于项目目录下，命名为XXX
    plt.clf()                                                                   #清空画板
    plt.plot(Several_independent_runs_fitness_list)                            #绘制若干次独立运行的最优路线长度折线图
    plt.savefig('SA + Tabu ('+str(independent_runs)+' times)')                                          #保存图片于项目目录下，命名为XXX

if __name__ == '__main__':
    main()

    
