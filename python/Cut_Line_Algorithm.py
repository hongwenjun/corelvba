import sys
# 从文件读取一组节点
nodes = []
with open(r"C:\TSP\CDR_TO_TSP", "r") as file:
    nodes = file.readlines()

# 删除第一个元素(源数据总数)
nodes = nodes[1:]
# 删除最后一个元素 换行(\n)
nodes = nodes[:-1]


# 定义一个函数来对节点进行排序和删除重复
def sort_and_remove_duplicates(nodes):
    sorted_nodes = sorted(nodes)  # 按照默认的元素顺序排序
    unique_nodes = [sorted_nodes[0]]  # 保留第一个元素

    for node in sorted_nodes[1:]:
        # 如果当前节点与前一个节点不相同，则将其添加到列表中
        if node != unique_nodes[-1]:
            unique_nodes.append(node)
    return unique_nodes



# 对节点进行排序和删除重复，写回文件
unique_nodes = sort_and_remove_duplicates(nodes)
total = len(unique_nodes)

points = []
for i in range(total):
    pot = unique_nodes[i].split()
    node = float(pot[0]), float(pot[1])
    points.append(node)

# print(points)

# boundary 是一个四元组，表示这组点的边界
x_min = min(point[0] for point in points)
x_max = max(point[0] for point in points)
y_min = min(point[1] for point in points)
y_max = max(point[1] for point in points)
boundary = (x_min, x_max, y_min, y_max)

print(boundary)

# near_boundary_points将会是一个列表，包含所有距离边界不超过1的点
near_boundary_points = []
for point in points:
    if (
        abs(point[0] - x_min) <= 1
        or abs(point[0] - x_max) <= 1
        or abs(point[1] - y_min) <= 1
        or abs(point[1] - y_max) <= 1
    ):
        near_boundary_points.append(point)

# print(near_boundary_points)

tuples_list = near_boundary_points
# 对元组列表分别按第二个元素进行升序排序 再第一个元素进行升序排序
sorted_tuples = sorted(tuples_list, key=lambda x: x[1])
near_boundary_points = sorted(sorted_tuples, key=lambda x: x[0])
# 输出排序后的元组列表
# print(near_boundary_points)

# 移除相邻点
def remove_adjacent(nodes):
    threshold = 0.5  # 设定阈值

    for i, node in enumerate(nodes):
        remove_indices = []
        for j, other_node in enumerate(nodes[i+1:]):
            distance = ((node[0]-other_node[0])**2 + (node[1]-other_node[1])**2)**0.5
            if distance < threshold:
                remove_indices.append(i+j+1)  # 记录需要移除的点的索引
        for index in sorted(remove_indices, reverse=True):
            del nodes[index]  # 移除邻近的点
    return nodes
near_boundary_points = remove_adjacent(near_boundary_points)

# 统计 x_min和y_max邻近的坐标点个数, 也就是求边界框左边 节点个数
left_points = []
top_points = []
for x, y in near_boundary_points:
    if abs(x - x_min) <= 0.5:
        left_points.append((x, y))

    if abs(y - y_max) <= 0.5:
        top_points.append((x, y))

left_points.reverse()
print(left_points)
if len(left_points) % 2 :
    top_points.reverse()    # 控制竖线从右往左
print(top_points)

# 把裁切线节点 写文件 TSP2.tx ，完成算法
total = len(left_points) + len(top_points)
f = open(r"C:\TSP\TSP2.txt", "w")
line = "%d %d\n" % (total, total)
f.write(line)

ext = 3  # extend 延长线, 默认值 3mm
if len(sys.argv) > 1:
    ext = float(sys.argv[1])
 
inverter = 1   # 交流频率控制
for x, y in left_points:
    if inverter == 1:
        line = "%f %f %f %f\n" % (x - ext, y, x_max + ext, y)
    else:
        line = "%f %f %f %f\n" % (x_max + ext, y, x - ext, y)

    f.write(line)
    inverter = (inverter + 1) % 2

for x, y in top_points:
    if inverter == len(left_points) % 2:  # 控制竖线从下面往上
        line = "%f %f %f %f\n" % (x, y + ext, x, y_min - ext)
    else:
        line = "%f %f %f %f\n" % (x, y_min - ext, x, y + ext)
    f.write(line)
    inverter = (inverter + 1) % 2
