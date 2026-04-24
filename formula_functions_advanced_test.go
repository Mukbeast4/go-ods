package goods

import (
	"math"
	"testing"
)

func asFloat(t *testing.T, v interface{}, err error) float64 {
	t.Helper()
	if err != nil {
		t.Fatalf("eval: %v", err)
	}
	f, ok := toFloat(v)
	if !ok {
		t.Fatalf("result %v not numeric", v)
	}
	return f
}

func mustFloat(t *testing.T, fn func() (interface{}, error)) float64 {
	t.Helper()
	v, err := fn()
	return asFloat(t, v, err)
}

func TestMEDIAN(t *testing.T) {
	cases := []struct {
		args []interface{}
		want float64
	}{
		{[]interface{}{1.0, 2.0, 3.0}, 2.0},
		{[]interface{}{1.0, 2.0, 3.0, 4.0}, 2.5},
		{[]interface{}{[]interface{}{10.0, 20.0, 30.0, 40.0, 50.0}}, 30.0},
	}
	for _, tc := range cases {
		got := mustFloat(t, func() (interface{}, error) { return evalMEDIAN(tc.args) })
		if got != tc.want {
			t.Errorf("MEDIAN(%v) = %v, want %v", tc.args, got, tc.want)
		}
	}
}

func TestMEDIANEmpty(t *testing.T) {
	if _, err := evalMEDIAN(nil); err == nil {
		t.Error("want error for empty MEDIAN")
	}
}

func TestSTDEVAndVAR(t *testing.T) {
	args := []interface{}{2.0, 4.0, 4.0, 4.0, 5.0, 5.0, 7.0, 9.0}

	sv := mustFloat(t, func() (interface{}, error) { return evalSTDEV(args) })
	if math.Abs(sv-2.138089935299) > 1e-6 {
		t.Errorf("STDEV = %v, want ~2.138", sv)
	}
	pv := mustFloat(t, func() (interface{}, error) { return evalSTDEVP(args) })
	if math.Abs(pv-2.0) > 1e-6 {
		t.Errorf("STDEVP = %v, want 2.0", pv)
	}

	vv := mustFloat(t, func() (interface{}, error) { return evalVAR(args) })
	if math.Abs(vv-4.571428571) > 1e-6 {
		t.Errorf("VAR = %v, want ~4.571", vv)
	}
	vpv := mustFloat(t, func() (interface{}, error) { return evalVARP(args) })
	if math.Abs(vpv-4.0) > 1e-6 {
		t.Errorf("VARP = %v, want 4.0", vpv)
	}
}

func TestRANK(t *testing.T) {
	ref := []interface{}{10.0, 20.0, 30.0, 20.0, 50.0}
	got := mustFloat(t, func() (interface{}, error) { return evalRANK([]interface{}{30.0, ref}) })
	if got != 2 {
		t.Errorf("RANK(30, ref) desc = %v, want 2", got)
	}
	got = mustFloat(t, func() (interface{}, error) { return evalRANK([]interface{}{30.0, ref, 1.0}) })
	if got != 4 {
		t.Errorf("RANK(30, ref) asc = %v, want 4", got)
	}
}

func TestLARGEAndSMALL(t *testing.T) {
	data := []interface{}{[]interface{}{3.0, 5.0, 7.0, 9.0, 11.0}}
	large2 := mustFloat(t, func() (interface{}, error) { return evalLARGE([]interface{}{data[0], 2.0}) })
	if large2 != 9 {
		t.Errorf("LARGE(data,2) = %v, want 9", large2)
	}
	small3 := mustFloat(t, func() (interface{}, error) { return evalSMALL([]interface{}{data[0], 3.0}) })
	if small3 != 7 {
		t.Errorf("SMALL(data,3) = %v, want 7", small3)
	}
}

func TestLARGEOutOfRange(t *testing.T) {
	data := []interface{}{[]interface{}{1.0, 2.0}, 5.0}
	if _, err := evalLARGE(data); err == nil {
		t.Error("want error when k exceeds dataset size")
	}
}

func TestPERCENTILE(t *testing.T) {
	nums := []interface{}{[]interface{}{1.0, 2.0, 3.0, 4.0, 5.0}}
	cases := []struct {
		p    float64
		want float64
	}{
		{0, 1}, {0.5, 3}, {1, 5}, {0.25, 2},
	}
	for _, tc := range cases {
		p := tc.p
		got := mustFloat(t, func() (interface{}, error) { return evalPERCENTILE([]interface{}{nums[0], p}) })
		if math.Abs(got-tc.want) > 1e-9 {
			t.Errorf("PERCENTILE p=%v = %v, want %v", tc.p, got, tc.want)
		}
	}
}

func TestRAND(t *testing.T) {
	seen := map[float64]bool{}
	for i := 0; i < 20; i++ {
		v, err := evalRAND(nil)
		if err != nil {
			t.Fatalf("RAND: %v", err)
		}
		f := v.(float64)
		if f < 0 || f >= 1 {
			t.Errorf("RAND = %v, want in [0,1)", f)
		}
		seen[f] = true
	}
	if len(seen) < 10 {
		t.Errorf("RAND: too many duplicates in 20 calls (got %d unique)", len(seen))
	}
}

func TestRANDBETWEEN(t *testing.T) {
	for i := 0; i < 50; i++ {
		v, err := evalRANDBETWEEN([]interface{}{5.0, 10.0})
		if err != nil {
			t.Fatalf("RANDBETWEEN: %v", err)
		}
		f := v.(float64)
		if f < 5 || f > 10 {
			t.Errorf("RANDBETWEEN = %v, want in [5,10]", f)
		}
		if f != math.Trunc(f) {
			t.Errorf("RANDBETWEEN = %v, want integer", f)
		}
	}
}

func TestRANDBETWEENInvalid(t *testing.T) {
	if _, err := evalRANDBETWEEN([]interface{}{10.0, 5.0}); err == nil {
		t.Error("want error when upper < lower")
	}
}

func TestIFS(t *testing.T) {
	cases := []struct {
		args []interface{}
		want interface{}
	}{
		{[]interface{}{true, "a", false, "b"}, "a"},
		{[]interface{}{false, "a", true, "b"}, "b"},
		{[]interface{}{false, "a", false, "b", true, "c"}, "c"},
	}
	for _, tc := range cases {
		got, err := evalIFS(tc.args)
		if err != nil {
			t.Errorf("IFS(%v): %v", tc.args, err)
			continue
		}
		if got != tc.want {
			t.Errorf("IFS(%v) = %v, want %v", tc.args, got, tc.want)
		}
	}
}

func TestIFSNoMatch(t *testing.T) {
	if _, err := evalIFS([]interface{}{false, "a", false, "b"}); err == nil {
		t.Error("want error when no condition matches")
	}
}

func TestSWITCH(t *testing.T) {
	cases := []struct {
		args []interface{}
		want interface{}
	}{
		{[]interface{}{2.0, 1.0, "one", 2.0, "two"}, "two"},
		{[]interface{}{5.0, 1.0, "one", 2.0, "two", "default"}, "default"},
		{[]interface{}{"a", "a", "matched"}, "matched"},
	}
	for _, tc := range cases {
		got, err := evalSWITCH(tc.args)
		if err != nil {
			t.Errorf("SWITCH(%v): %v", tc.args, err)
			continue
		}
		if got != tc.want {
			t.Errorf("SWITCH(%v) = %v, want %v", tc.args, got, tc.want)
		}
	}
}

func TestXLOOKUP(t *testing.T) {
	keys := []interface{}{"a", "b", "c"}
	vals := []interface{}{1.0, 2.0, 3.0}

	got, err := evalXLOOKUP([]interface{}{"b", keys, vals})
	if err != nil {
		t.Fatalf("XLOOKUP: %v", err)
	}
	if got != 2.0 {
		t.Errorf("XLOOKUP(b) = %v, want 2", got)
	}
}

func TestXLOOKUPDefault(t *testing.T) {
	keys := []interface{}{"a", "b"}
	vals := []interface{}{1.0, 2.0}
	got, err := evalXLOOKUP([]interface{}{"z", keys, vals, "missing"})
	if err != nil {
		t.Fatalf("XLOOKUP: %v", err)
	}
	if got != "missing" {
		t.Errorf("XLOOKUP default = %v, want missing", got)
	}
}

func TestXLOOKUPMismatchedArrays(t *testing.T) {
	_, err := evalXLOOKUP([]interface{}{"a", []interface{}{"a", "b"}, []interface{}{1.0}})
	if err == nil {
		t.Error("want error for arrays of different lengths")
	}
}

func TestFunctionsRegisteredInMap(t *testing.T) {
	names := []string{
		"MEDIAN", "STDEV", "STDEVP", "VAR", "VARP",
		"RANK", "LARGE", "SMALL", "PERCENTILE",
		"RAND", "RANDBETWEEN",
		"IFS", "SWITCH", "XLOOKUP",
	}
	for _, name := range names {
		if _, ok := formulaFunctions[name]; !ok {
			t.Errorf("function %s not registered", name)
		}
	}
}

func TestEvaluateMEDIANViaFormula(t *testing.T) {
	got, err := Evaluate("MEDIAN(1;2;3;4;5)", nil)
	if err != nil {
		t.Fatalf("Evaluate: %v", err)
	}
	if got != 3.0 {
		t.Errorf("Evaluate MEDIAN = %v, want 3", got)
	}
}

func TestEvaluateXLOOKUPViaFormula(t *testing.T) {
	f := NewFile()
	f.SetCellValue("Sheet1", "A1", "alpha")
	f.SetCellValue("Sheet1", "A2", "beta")
	f.SetCellValue("Sheet1", "A3", "gamma")
	f.SetCellFloat("Sheet1", "B1", 10)
	f.SetCellFloat("Sheet1", "B2", 20)
	f.SetCellFloat("Sheet1", "B3", 30)
	f.SetCellFormula("Sheet1", "C1", `XLOOKUP("beta";[.A1:A3];[.B1:B3])`)

	got, err := f.EvaluateFormula("Sheet1", "C1", nil)
	if err != nil {
		t.Fatalf("EvaluateFormula: %v", err)
	}
	if got != 20.0 {
		t.Errorf("XLOOKUP(beta) via formula = %v, want 20", got)
	}
}
