package util

import (
	"testing"
)

func TestProcess(t *testing.T) {
	result, err := Process("1+2*3")
	if err != nil {
		t.Fatal(err)
	}
	if result != 7 {
		t.Error("calc err!")
	}
	result, err = Process("2*3+5")
	if err != nil {
		t.Fatal(err)
	}
	if result != 11 {
		t.Error("calc err!")
	}
	result, err = Process("100")
	if err != nil {
		t.Fatal(err)
	}
	if result != 100 {
		t.Error("calc err!")
	}
	result, err = Process("100*(1+2*2)")
	if err != nil {
		t.Fatal(err)
	}
	if result != 500 {
		t.Error("calc err!")
	}
	t.Log("ok , success")
}
