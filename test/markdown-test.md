# Markdown 语法测试

## 标题语法
# 标题 H1
## 标题 H2
### 标题 H3
#### 标题 H4
##### 标题 H5
###### 标题 H6

标题 H1 (下划线式)
================

标题 H2 (下划线式)
----------------


## 强调语法

*斜体*

_斜体_

**粗体**

__粗体__

***粗斜体***

___粗斜体___

~~删除线~~


## 列表语法


* 无序列表
+ 无序列表
- 无序列表


1. 有序列表
2. 有序列表
3. 有序列表
4. 有序列表
5. 有序列表
6. 有序列表
7. 有序列表
8. 有序列表
9. 有序列表
10. 有序列表
11. 有序列表
12. 有序列表
13. 有序列表
14. 有序列表
15. 有序列表


- [x] 任务列表
- [ ] 任务列表

## 引用语法

> 引用

>> 嵌套引用


## 代码语法

### 行内代码

`pandocPath = Path.Combine(programFiles, "Pandoc", "pandoc.exe");`


### 代码块

```

# 正则表达式示例
import re
pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'

```

### 代码块 (带C语言标识)

```c

#include <stdio.h>

// 计算斐波那契数列
int fibonacci(int n) {
    if (n <= 1) {
        return n;
    }
    return fibonacci(n-1) + fibonacci(n-2);
}

int main() {
    int n = 10;
    printf("斐波那契数列前 %d 项:\n", n);
    for (int i = 0; i < n; i++) {
        printf("%d ", fibonacci(i));
    }
    printf("\n");
    return 0;
}

```

### 代码块 (带Python语言标识)

```python

def fibonacci(n):
    """计算斐波那契数列"""
    if n <= 1:
        return n
    return fibonacci(n-1) + fibonacci(n-2)

if __name__ == "__main__":
    n = 10
    print(f"斐波那契数列前 {n} 项:")
    for i in range(n):
        print(fibonacci(i), end=" ")
    print()

```


## 链接语法

[行内链接](https://example.com)

[参考链接][1]

<https://example.com>

https://example.com

[1]: https://example.com


## 图片语法

### 行内图片

![图片1标题](https://images.unsplash.com/photo-1506744038136-46273834b3fb?w=600)

### 参考图片

![图片2标题][2]

[2]: https://images.unsplash.com/photo-1506744038136-46273834b3fb?w=600


## 表格语法

| 表头1 | 表头2 | 表头3 | 表头4 | 表头5 | 表头6 |
|-------|-------|-------|-------|-------|-------|
| 单元格1 | 单元格2 | 单元格3 | 单元格4 | 单元格5 | 单元格6 |
| Cell 1 | Cell 2 | Cell 3 | Cell 4 | Cell 5 | Cell 6 |
| 单元Cell 1 | 单元Cell 2 | 单元Cell 3 | 单元Cell 4 | 单元Cell 5 | 超超超超级长单元Cell 6 |


## 水平线语法

---
***
___


## 脚注语法

这是一个脚注引用[^1]

[^1]: 这是脚注定义


## 转义字符

\*星号\*  
\_下划线\_  
\`反引号\`  
\#井号\#  
\+加号\+  
\-减号\-  
\.点号\.  
\!感叹号\!


## 公式

行内公式示例

$E=mc^2$

Latex公式1

$$
\begin{aligned}
f(x) &= \int_{-\infty}^\infty \hat f(\xi)\,e^{2 \pi i \xi x} \,d\xi \\
\frac{\partial u}{\partial t} &= H(u) + \alpha \left( \frac{\partial^2 u}{\partial x^2} + \frac{\partial^2 u}{\partial y^2} \right) \\
\begin{pmatrix}
a & b \\
c & d \\
\end{pmatrix}
\times
\begin{pmatrix}
x \\
y \\
\end{pmatrix}
&=
\begin{pmatrix}
ax + by \\
cx + dy \\
\end{pmatrix}
\end{aligned}
$$

Latex公式2

$$
\sum_{n=1}^\infty \frac{1}{n^2} = \frac{\pi^2}{6}
$$


