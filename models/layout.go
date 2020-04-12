package models

type Layout struct {
	Left   int
	Top    int
	Right  int
	Bottom int
}

var DefaultLayout Layout = Layout{
	Left:   1,
	Top:    1,
	Right:  1,
	Bottom: 1,
}