package util

import (
	"errors"
	"fmt"
	"github.com/bluebuff/simple-excelize/pkg/stack"
	"strconv"
	"strings"
)

//	输入一个四则运算表达式（中缀表达式），求出其结果。如：9+(3-1)*3+10/2，结果为 20。
func Process(express string) (int64, error) {
	if len(express) == 0 {
		return 0, errors.New("invalid express")
	}

	//	将中缀表达式转换成后缀表达式（逆波兰式），postfixExpress：后缀表达式
	express = strings.TrimSpace(express)
	postfixExpress, err := transPostfixExpress(express)
	if err != nil {
		return 0, err
	}

	//	后缀表达式求值
	return calc(postfixExpress)
}

//	将中缀表达式转换成后缀表达式（逆波兰式），postfixExpress：后缀表达式
func transPostfixExpress(express string) (postfixExpress []string, err error) {
	var (
		opStack stack.Stack //	运算符堆栈
		i       int
	)

LABEL:
	for i < len(express) { //	从左至右扫描中缀表达式
		switch {
		//	1. 若读取的是操作数，则将该操作数存入后缀表达式。
		case express[i] >= '0' && express[i] <= '9':
			var number []byte //	如数字123，由'1'、'2'、'3'组成
			for ; i < len(express); i++ {
				if express[i] < '0' || express[i] > '9' {
					break
				}
				number = append(number, express[i])
			}
			postfixExpress = append(postfixExpress, string(number))

			//	2. 若读取的是运算符：
			//	(1) 该运算符为左括号"("，则直接压入运算符堆栈。
		case express[i] == '(':
			opStack.Push(fmt.Sprintf("%c", express[i]))
			i++

			//	(2) 该运算符为右括号")"，则输出运算符堆栈中的运算符到后缀表达式，直到遇到左括号为止。
		case express[i] == ')':
			for !opStack.IsEmpty() {
				data, _ := opStack.Pop()
				if data[0] == '(' {
					break
				}
				postfixExpress = append(postfixExpress, data)
			}
			i++

			//	(3) 该运算符为非括号运算符:
		case express[i] == '+' || express[i] == '-' || express[i] == '*' || express[i] == '/':
			//	(a)若运算符堆栈为空,则直接压入运算符堆栈。
			if opStack.IsEmpty() {
				opStack.Push(fmt.Sprintf("%c", express[i]))
				i++
				continue LABEL
			}

			data, _ := opStack.Top()
			//	(b)若运算符堆栈栈顶的运算符为括号，则直接压入运算符堆栈。(只可能为左括号这种情况)
			if data[0] == '(' {
				opStack.Push(fmt.Sprintf("%c", express[i]))
				i++
				continue LABEL
			}
			//	(c)若比运算符堆栈栈顶的运算符优先级低或相等，则输出栈顶运算符到后缀表达式,并将当前运算符压入运算符堆栈。
			if (express[i] == '+' || express[i] == '-') ||
				((express[i] == '*' || express[i] == '/') && (data[0] == '*' || data[0] == '/')) {
				postfixExpress = append(postfixExpress, data)
				opStack.Pop()
				opStack.Push(fmt.Sprintf("%c", express[i]))
				i++
				continue LABEL
			}
			//	(d)若比运算符堆栈栈顶的运算符优先级高，则直接压入运算符堆栈。
			opStack.Push(fmt.Sprintf("%c", express[i]))
			i++

		default:
			err = fmt.Errorf("invalid express:%v", express[i])
			return
		}
	}

	//	3. 扫描结束，将运算符堆栈中的运算符依次弹出，存入后缀表达式。
	for !opStack.IsEmpty() {
		data, _ := opStack.Pop()
		if data[0] == '#' {
			break
		}
		postfixExpress = append(postfixExpress, data)
	}

	return
}

//	后缀表达式求值
func calc(postfixExpress []string) (result int64, err error) {
	var (
		num1 string
		num2 string
		s    stack.Stack //	操作栈，用于存入操作数，运算符
	)

	//	从左至右扫描后缀表达式
	for i := 0; i < len(postfixExpress); i++ {
		var cur = postfixExpress[i]

		//	1. 若读取的是运算符
		if cur[0] == '+' || cur[0] == '-' || cur[0] == '*' || cur[0] == '/' {
			//	从操作栈中弹出两个数进行运算
			num1, err = s.Pop()
			if err != nil {
				return
			}
			num2, err = s.Pop()
			if err != nil {
				return
			}

			//	先弹出的数为B，后弹出的数为A
			B, _ := strconv.Atoi(num1)
			A, _ := strconv.Atoi(num2)
			var res int

			switch cur[0] {
			case '+':
				res = A + B
			case '-':
				res = A - B
			case '*':
				res = A * B
			case '/':
				res = A / B
			default:
				err = fmt.Errorf("invalid operation")
				return
			}

			//	将中间结果压栈
			s.Push(fmt.Sprintf("%d", res))
			//fmt.Println("mid value = ", res)
		} else {
			//	1. 若读取的是操作数，直接压栈
			s.Push(cur)
		}
	}

	//	计算结束，栈顶保存最后结果
	resultStr, err := s.Top()
	if err != nil {
		return
	}
	result, err = strconv.ParseInt(resultStr, 10, 64)
	return
}
