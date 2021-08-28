package util

import (
	"strconv"
	"strings"
)

func MustInt(data string) int {
	data = strings.TrimSpace(data)
	num, _ := strconv.ParseInt(data, 10, 32)
	return int(num)
}

func MustInt32(data string) int32 {
	data = strings.TrimSpace(data)
	num, _ := strconv.ParseInt(data, 10, 32)
	return int32(num)
}

func MustInt64(data string) int64 {
	data = strings.TrimSpace(data)
	num, _ := strconv.ParseInt(data, 10, 64)
	return num
}

func MustUint32(data string) uint32 {
	data = strings.TrimSpace(data)
	num, _ := strconv.ParseUint(data, 10, 32)
	return uint32(num)
}

func MustUint64(data string) uint64 {
	data = strings.TrimSpace(data)
	num, _ := strconv.ParseUint(data, 10, 64)
	return num
}

func MustFloat32(data string) float32 {
	data = strings.TrimSpace(data)
	num, _ := strconv.ParseFloat(data, 32)
	return float32(num)
}

func MustFloat64(data string) float64 {
	data = strings.TrimSpace(data)
	num, _ := strconv.ParseFloat(data, 64)
	return num
}