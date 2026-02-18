# TODO

Track unrelated issues noticed while reviewing PRs.

## New Issues
- [ ] Optimize `ValidateRoList.vbs` status detection so each RO check evaluates both outcomes (`NOT ON FILE` and `(PFC) POST FINAL CHARGES`) in the same polling loop; avoid waiting a full timeout when one condition is already visible.
	- Immediate ideas:
		- Check both target strings on every poll from the same screen buffer (single read, single decision).
		- Add an early-exit path for known terminal/error states that clearly mean neither target will appear.
		- Use a short/fast timeout for the first detection window, then optional longer fallback only when screen is still transitioning.
		- Keep per-poll logging lightweight (only on match, timeout, or every N polls) to reduce overhead.
		- Preserve existing fail-fast behavior and output contract (`RO,STATUS`) to avoid downstream breakage.

## Follow-up (Non-blocking)
- [ ] 

## Nice-to-have
- [ ] 
