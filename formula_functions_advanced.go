package goods

import (
	"fmt"
	"math"
	"math/rand/v2"
	"sort"
)

func collectNumbers(args []interface{}) []float64 {
	var out []float64
	for _, a := range args {
		if vals := flattenArg(a); vals != nil {
			for _, v := range vals {
				if isEmptyValue(v) {
					continue
				}
				if f, ok := toFloat(v); ok {
					out = append(out, f)
				}
			}
		} else {
			if isEmptyValue(a) {
				continue
			}
			if f, ok := toFloat(a); ok {
				out = append(out, f)
			}
		}
	}
	return out
}

func evalMEDIAN(args []interface{}) (interface{}, error) {
	nums := collectNumbers(args)
	if len(nums) == 0 {
		return nil, fmt.Errorf("MEDIAN: no numeric values")
	}
	sort.Float64s(nums)
	n := len(nums)
	if n%2 == 1 {
		return nums[n/2], nil
	}
	return (nums[n/2-1] + nums[n/2]) / 2, nil
}

func variance(nums []float64, sample bool) (float64, error) {
	if len(nums) == 0 || (sample && len(nums) < 2) {
		return 0, fmt.Errorf("not enough data points")
	}
	var mean float64
	for _, v := range nums {
		mean += v
	}
	mean /= float64(len(nums))
	var sum float64
	for _, v := range nums {
		d := v - mean
		sum += d * d
	}
	denom := float64(len(nums))
	if sample {
		denom = float64(len(nums) - 1)
	}
	return sum / denom, nil
}

func evalVAR(args []interface{}) (interface{}, error) {
	nums := collectNumbers(args)
	return variance(nums, true)
}

func evalVARP(args []interface{}) (interface{}, error) {
	nums := collectNumbers(args)
	return variance(nums, false)
}

func evalSTDEV(args []interface{}) (interface{}, error) {
	v, err := evalVAR(args)
	if err != nil {
		return nil, err
	}
	return math.Sqrt(v.(float64)), nil
}

func evalSTDEVP(args []interface{}) (interface{}, error) {
	v, err := evalVARP(args)
	if err != nil {
		return nil, err
	}
	return math.Sqrt(v.(float64)), nil
}

func evalRANK(args []interface{}) (interface{}, error) {
	if len(args) < 2 {
		return nil, fmt.Errorf("RANK requires at least 2 arguments")
	}
	target, ok := toFloat(args[0])
	if !ok {
		return nil, fmt.Errorf("RANK: number must be numeric")
	}
	nums := collectNumbers(args[1:2])
	if len(nums) == 0 {
		return nil, fmt.Errorf("RANK: empty reference")
	}

	ascending := false
	if len(args) >= 3 {
		if f, ok := toFloat(args[2]); ok && f != 0 {
			ascending = true
		}
	}

	sorted := make([]float64, len(nums))
	copy(sorted, nums)
	if ascending {
		sort.Float64s(sorted)
	} else {
		sort.Sort(sort.Reverse(sort.Float64Slice(sorted)))
	}

	for i, v := range sorted {
		if v == target {
			return float64(i + 1), nil
		}
	}
	return nil, fmt.Errorf("RANK: value not found in reference")
}

func evalLARGE(args []interface{}) (interface{}, error) {
	if len(args) < 2 {
		return nil, fmt.Errorf("LARGE requires 2 arguments")
	}
	k, ok := toFloat(args[1])
	if !ok || k < 1 {
		return nil, fmt.Errorf("LARGE: k must be >= 1")
	}
	nums := collectNumbers(args[:1])
	if int(k) > len(nums) {
		return nil, fmt.Errorf("LARGE: k exceeds dataset size")
	}
	sort.Sort(sort.Reverse(sort.Float64Slice(nums)))
	return nums[int(k)-1], nil
}

func evalSMALL(args []interface{}) (interface{}, error) {
	if len(args) < 2 {
		return nil, fmt.Errorf("SMALL requires 2 arguments")
	}
	k, ok := toFloat(args[1])
	if !ok || k < 1 {
		return nil, fmt.Errorf("SMALL: k must be >= 1")
	}
	nums := collectNumbers(args[:1])
	if int(k) > len(nums) {
		return nil, fmt.Errorf("SMALL: k exceeds dataset size")
	}
	sort.Float64s(nums)
	return nums[int(k)-1], nil
}

func evalPERCENTILE(args []interface{}) (interface{}, error) {
	if len(args) < 2 {
		return nil, fmt.Errorf("PERCENTILE requires 2 arguments")
	}
	p, ok := toFloat(args[1])
	if !ok || p < 0 || p > 1 {
		return nil, fmt.Errorf("PERCENTILE: k must be between 0 and 1")
	}
	nums := collectNumbers(args[:1])
	if len(nums) == 0 {
		return nil, fmt.Errorf("PERCENTILE: empty dataset")
	}
	sort.Float64s(nums)
	if len(nums) == 1 {
		return nums[0], nil
	}

	pos := p * float64(len(nums)-1)
	lower := int(math.Floor(pos))
	upper := int(math.Ceil(pos))
	if lower == upper {
		return nums[lower], nil
	}
	frac := pos - float64(lower)
	return nums[lower] + frac*(nums[upper]-nums[lower]), nil
}

func evalRAND(args []interface{}) (interface{}, error) {
	if len(args) != 0 {
		return nil, fmt.Errorf("RAND takes no arguments")
	}
	return rand.Float64(), nil
}

func evalRANDBETWEEN(args []interface{}) (interface{}, error) {
	if len(args) != 2 {
		return nil, fmt.Errorf("RANDBETWEEN requires 2 arguments")
	}
	lo, lok := toFloat(args[0])
	hi, hok := toFloat(args[1])
	if !lok || !hok {
		return nil, fmt.Errorf("RANDBETWEEN: bounds must be numeric")
	}
	if hi < lo {
		return nil, fmt.Errorf("RANDBETWEEN: upper bound below lower")
	}
	loInt := int64(math.Ceil(lo))
	hiInt := int64(math.Floor(hi))
	if hiInt < loInt {
		return nil, fmt.Errorf("RANDBETWEEN: no integer in range")
	}
	return float64(loInt + rand.Int64N(hiInt-loInt+1)), nil
}

func evalIFS(args []interface{}) (interface{}, error) {
	if len(args) < 2 || len(args)%2 != 0 {
		return nil, fmt.Errorf("IFS requires pairs of condition/value")
	}
	for i := 0; i < len(args); i += 2 {
		if toBool(args[i]) {
			return args[i+1], nil
		}
	}
	return nil, fmt.Errorf("IFS: no condition matched")
}

func evalSWITCH(args []interface{}) (interface{}, error) {
	if len(args) < 3 {
		return nil, fmt.Errorf("SWITCH requires at least 3 arguments")
	}
	key := args[0]
	i := 1
	for i+1 < len(args) {
		if valuesEqual(key, args[i]) {
			return args[i+1], nil
		}
		i += 2
	}
	if i < len(args) {
		return args[i], nil
	}
	return nil, fmt.Errorf("SWITCH: no match and no default")
}

func evalXLOOKUP(args []interface{}) (interface{}, error) {
	if len(args) < 3 {
		return nil, fmt.Errorf("XLOOKUP requires at least 3 arguments")
	}
	lookup := args[0]
	lookupArray := flattenArg(args[1])
	returnArray := flattenArg(args[2])
	if lookupArray == nil || returnArray == nil {
		return nil, fmt.Errorf("XLOOKUP: lookup_array and return_array must be ranges")
	}
	if len(lookupArray) != len(returnArray) {
		return nil, fmt.Errorf("XLOOKUP: arrays must have matching lengths")
	}

	for i, v := range lookupArray {
		if valuesEqual(lookup, v) {
			return returnArray[i], nil
		}
	}
	if len(args) >= 4 {
		return args[3], nil
	}
	return nil, fmt.Errorf("XLOOKUP: value not found")
}

func valuesEqual(a, b interface{}) bool {
	if af, aok := toFloat(a); aok {
		if bf, bok := toFloat(b); bok {
			return af == bf
		}
	}
	as := fmt.Sprintf("%v", a)
	bs := fmt.Sprintf("%v", b)
	return as == bs
}

func init() {
	additional := map[string]formulaFunc{
		"MEDIAN":      evalMEDIAN,
		"STDEV":       evalSTDEV,
		"STDEVP":      evalSTDEVP,
		"VAR":         evalVAR,
		"VARP":        evalVARP,
		"RANK":        evalRANK,
		"LARGE":       evalLARGE,
		"SMALL":       evalSMALL,
		"PERCENTILE":  evalPERCENTILE,
		"RAND":        evalRAND,
		"RANDBETWEEN": evalRANDBETWEEN,
		"IFS":         evalIFS,
		"SWITCH":      evalSWITCH,
		"XLOOKUP":     evalXLOOKUP,
	}
	for k, v := range additional {
		formulaFunctions[k] = v
	}
}
