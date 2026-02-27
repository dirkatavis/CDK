# PostFinalCharges TODO

## Critical Bugs
- [ ] Investigate `Variable is undefined: 'DoEvents'` error appearing in CLI test mode. WSH environment does not support native `DoEvents`. Consider swapping with `WScript.Sleep 1` for yielding or defining a no-op global if UI-less.

## Features
- [ ] Implement multi-RO batch processing from `CashoutRoList.csv`.

## Performance & Stability
- [ ] Optimize `WaitForTextAtBottom` for high-latency scenarios (see `run_stress_tests.vbs`).
