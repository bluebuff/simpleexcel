package stack

import "fmt"

type Stack struct {
	data [1024]string
	top int
}

func (s *Stack) IsEmpty() bool {
	return s.top == 0
}

func (s *Stack) Top() (ret string, err error) {
	if s.top == 0 {
		err = fmt.Errorf("stack is empty")
		return
	}
	ret = s.data[s.top-1]
	return
}

func (s *Stack) Push(str string) {
	s.data[s.top] = str
	s.top++
}

func (s *Stack) Pop() (ret string, err error) {
	if s.top == 0 {
		err = fmt.Errorf("stack is empty")
		return
	}
	s.top--
	ret = s.data[s.top]
	return
}