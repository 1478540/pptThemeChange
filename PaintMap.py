

class PaintMap:
    map=[]
    shapes_color=[]

    hasAnswer=True
    ticket=0

    def __init__(self, map):
        self.map = map
        self.shapes_color = [0] * len(self.map)


    def paintMap(self):
        k = 0
        color = 1
        area = 0
        flag = 0

        self.shapes_color[area] = color  # 第一个区域先着色1
        area += 1

        # 给所有节点着色
        while area < len(self.map):
            color = self.shapes_color[area] + 1  # 从当前颜色的下一个颜色开始试

            flag = 0
            # 尝试为area着色
            while color <= 4 and flag == 0:

                k=0
                while k<area:
                    # 如果当前结点和所有与它相邻结点有重色的
                    if self.map[area][k] * self.shapes_color[k] == color:
                        color += 1
                        break  # 直接跳出循环，试下一种颜色
                    k += 1

                # 找到不重色的
                if k == area:
                    flag = 1

            # 如果怎么着色都会重色
            if color > 4:
                self.shapes_color[area] = 0  # 当前区域置0 ！！！注意这一步别忘掉
                area -= 1  # 前面一个重新着色
            else:
                self.shapes_color[area] = color
                area += 1


    def backCorrectColorScheme(self):
        self.paintMap()
        return self.shapes_color




