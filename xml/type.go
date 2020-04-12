package xml

type Type string

const (
	Default Type = ""
	String  Type = "string"
	Uint32  Type = "uint32"
	Uint64  Type = "unit64"
	Float32 Type = "float32"
	Float64 Type = "float64"
	Title   Type = "title"
	Header  Type = "header"
	Formula Type = "formula"
)