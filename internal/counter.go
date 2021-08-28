package internal

import "sync"

type Counter interface {
	Incr() int
	Current() int
	IncrBy(offset int) int
	Reset()
}

func NewCounter(base int) Counter {
	counter := smartCounter{index: base, base: base}
	return &counter
}

type smartCounter struct {
	base  int
	index int
	lock  sync.RWMutex
}

func (counter *smartCounter) Incr() int {
	defer func() {
		counter.lock.Lock()
		counter.index++
		counter.lock.Unlock()
	}()
	return counter.Current()
}

func (counter *smartCounter) Current() int {
	counter.lock.RLock()
	defer counter.lock.RUnlock()
	return counter.index
}

func (counter *smartCounter) IncrBy(offset int) int {
	defer func() {
		counter.lock.Lock()
		counter.index += offset
		counter.lock.Unlock()
	}()
	return counter.Current()
}

func (counter *smartCounter) Reset() {
	counter.lock.Lock()
	counter.index = counter.base
	counter.lock.Unlock()
}
