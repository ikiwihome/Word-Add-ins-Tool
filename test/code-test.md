# 多语言代码示例

## C# 示例

```csharp
// Program.cs
using System;

namespace Example
{
    /// <summary>
    /// 主程序类
    /// </summary>
    class Program
    {
        /// <summary>
        /// 主入口点
        /// </summary>
        static void Main(string[] args)
        {
            Console.WriteLine("Hello, C# World!");
            int result = Add(5, 3);
            Console.WriteLine($"5 + 3 = {result}");
        }

        /// <summary>
        /// 两数相加
        /// </summary>
        /// <param name="a">第一个数</param>
        /// <param name="b">第二个数</param>
        /// <returns>相加结果</returns>
        static int Add(int a, int b)
        {
            return a + b;
        }
    }
}
```

## C++ 示例

```cpp

// main.cpp
#include <iostream>

/**
 * 数学工具类
 */
class MathUtils {
public:
    /**
     * 计算阶乘
     * @param n 输入整数
     * @return n的阶乘
     */
    static int factorial(int n) {
        if (n <= 1) return 1;
        return n * factorial(n - 1);
    }
};

int main() {
    std::cout << "C++ Factorial Example" << std::endl;
    int num = 5;
    std::cout << "Factorial of " << num << " is " 
              << MathUtils::factorial(num) << std::endl;
    return 0;
}
```

## Python 示例

```python
# example.py
"""这是一个Python示例模块"""

class Calculator:
    """简单的计算器类"""
    
    def __init__(self, name: str):
        """初始化计算器"""
        self.name = name
    
    def add(self, a: float, b: float) -> float:
        """两数相加"""
        return a + b
    
    def greet(self) -> str:
        """返回欢迎消息"""
        return f"Hello from {self.name} calculator!"

if __name__ == "__main__":
    calc = Calculator("Scientific")
    print(calc.greet())
    print(f"2.5 + 3.5 = {calc.add(2.5, 3.5)}")
```

## Shell 示例

```bash
#!/bin/bash
# system_info.sh

# 显示系统信息的脚本

# 主函数
main() {
    echo "System Information"
    echo "------------------"
    show_date
    show_uptime
    show_disk_usage
}

# 显示当前日期
show_date() {
    echo "Date: $(date '+%Y-%m-%d %H:%M:%S')"
}

# 显示系统运行时间
show_uptime() {
    echo "Uptime: $(uptime -p)"
}

# 显示磁盘使用情况
show_disk_usage() {
    echo "Disk Usage:"
    df -h | grep -v "tmpfs"
}

main "$@"
```

## Java 示例

```java
// Main.java
package com.example;

/**
 * Java示例程序
 */
public class Main {
    /**
     * 主方法
     * @param args 命令行参数
     */
    public static void main(String[] args) {
        System.out.println("Java Sample Program");
        
        Person person = new Person("Alice", 30);
        person.greet();
        
        System.out.println("Next year " + person.getName() + 
                          " will be " + person.getAgeNextYear());
    }
}

/**
 * 人类
 */
class Person {
    private String name;
    private int age;
    
    /**
     * 构造函数
     * @param name 姓名
     * @param age 年龄
     */
    public Person(String name, int age) {
        this.name = name;
        this.age = age;
    }
    
    /**
     * 问候方法
     */
    public void greet() {
        System.out.println("Hello, my name is " + name);
    }
    
    /**
     * 计算明年年龄
     * @return 明年年龄
     */
    public int getAgeNextYear() {
        return age + 1;
    }
    
    // Getter方法
    public String getName() {
        return name;
    }
}
```

## Javascript 示例

```javascript
// app.js
/**
 * 购物车类
 */
class ShoppingCart {
    constructor() {
        this.items = [];
    }

    /**
     * 添加商品
     * @param {string} name 商品名称
     * @param {number} price 商品价格
     * @param {number} quantity 商品数量
     */
    addItem(name, price, quantity = 1) {
        this.items.push({ name, price, quantity });
    }

    /**
     * 计算总价
     * @returns {number} 总金额
     */
    calculateTotal() {
        return this.items.reduce(
            (total, item) => total + (item.price * item.quantity), 0
        );
    }
}

// 使用示例
const cart = new ShoppingCart();
cart.addItem("Book", 29.99);
cart.addItem("Pen", 3.50, 5);

console.log(`Total: $${cart.calculateTotal().toFixed(2)}`);
```