# Performance Optimization Quick Reference

## Overview
This document provides a quick reference for the performance optimizations made to reduce hardcoded delays across the CDK automation scripts.

## Before vs After Comparison

### Delay Constants Optimized

#### CreateNew_ROs/Create_ROs.vbs
```vbscript
' BEFORE
Const POLL_INTERVAL = 1000        ' 1 second polling
Const POST_ENTRY_WAIT = 200       ' 200ms after entry
Const PRE_KEY_WAIT = 150          ' 150ms before keys
Const POST_KEY_WAIT = 350         ' 350ms after keys

' AFTER
Const POLL_INTERVAL = 500         ' 500ms polling (2x faster)
Const POST_ENTRY_WAIT = 100       ' 100ms after entry (50% faster)
Const PRE_KEY_WAIT = 100          ' 100ms before keys (33% faster)
Const POST_KEY_WAIT = 200         ' 200ms after keys (43% faster)
```

#### Close_ROs Files
```vbscript
' BEFORE
Const POST_ENTRY_WAIT = 200       ' 200ms
Const PRE_KEY_WAIT = 150          ' 150ms
Const POST_KEY_WAIT = 350         ' 350ms
Const DelayTimeAfterPromptDetection = 500  ' 500ms

' AFTER
Const POST_ENTRY_WAIT = 100       ' 100ms (50% faster)
Const PRE_KEY_WAIT = 100          ' 100ms (33% faster)
Const POST_KEY_WAIT = 200         ' 200ms (43% faster)
Const DelayTimeAfterPromptDetection = 200  ' 200ms (60% faster)
```

### Specific Delay Reductions

#### PostFinalCharges/PostFinalCharges.vbs
| Operation | Before | After | Savings |
|-----------|--------|-------|---------|
| IsStatusReady() | 1000ms | 200ms | 800ms (80%) |
| ADD A LABOR OPERATION | 2000ms | 800ms | 1200ms (60%) |
| SOLD HOURS | 1500ms | 800ms | 700ms (47%) |
| General response | 800ms | 500ms | 300ms (38%) |
| ALL LABOR POSTED pre-wait | 3000ms | 1000ms | 2000ms (67%) |
| ProcessOpenStatusLines | 1000ms | 500ms | 500ms (50%) |

#### Close_ROs/Close_ROs_Pt1.vbs
| Operation | Before | After | Savings |
|-----------|--------|-------|---------|
| Pre-CCC command | 1000ms | 300ms | 700ms (70%) |
| Post-CCC command | 2000ms | 500ms | 1500ms (75%) |
| Final exit | 1000ms | 500ms | 500ms (50%) |
| Screen update check | 2000ms | 500ms | 1500ms (75%) |
| Debug display | 2000ms | 200ms | 1800ms (90%) |

#### Close_ROs/Close_ROs_Pt2.vbs
| Operation | Before | After | Savings |
|-----------|--------|-------|---------|
| CheckForTextInLine2 | 2000ms | 300ms | 1700ms (85%) |
| INVOICE PRINTER | 2000ms | 500ms | 1500ms (75%) |
| WaitForTextAtBottom polling | 500ms | 250ms | 250ms (50%) |
| Text registration | 100ms | 50ms | 50ms (50%) |
| Post-enter wait | 500ms | 200ms | 300ms (60%) |
| All closeout steps | 1000ms | 500ms | 500ms (50%) |

## Cumulative Impact Example

### Single RO Processing (Estimated)
```
BEFORE:
- Prompt detection: 10 prompts × 1000ms = 10,000ms
- Screen transitions: 8 × 1500ms = 12,000ms
- Closeout sequence: 6 × 1000ms = 6,000ms
- Other delays: ~5,000ms
TOTAL: ~33 seconds per RO

AFTER:
- Prompt detection: 10 prompts × 500ms = 5,000ms (50% faster)
- Screen transitions: 8 × 500ms = 4,000ms (67% faster)
- Closeout sequence: 6 × 500ms = 3,000ms (50% faster)
- Other delays: ~2,000ms (60% faster)
TOTAL: ~14 seconds per RO

IMPROVEMENT: 19 seconds saved (58% faster)
```

### Batch Processing (100 ROs)
```
BEFORE: 100 ROs × 33 seconds = 3,300 seconds (~55 minutes)
AFTER:  100 ROs × 14 seconds = 1,400 seconds (~23 minutes)

TIME SAVED: 1,900 seconds (~32 minutes) for 100 ROs
```

## Key Principles Applied

### 1. Remove Redundancy
**Example:** Removed `bzhao.Pause(1000)` before key commands because `FastKey()` already handles timing.

### 2. Minimize Required Delays
**Example:** Reduced text registration from 100ms to 50ms - still sufficient for keyboard input processing.

### 3. Optimize Polling
**Example:** Reduced polling interval from 1000ms to 500ms - faster detection with minimal overhead.

### 4. Preserve Reliability
**Example:** Kept 800ms for ADD A LABOR OPERATION prompt due to screen mode change requirements.

## Configuration Constants Reference

### Recommended Minimum Values
Based on testing and analysis:

```vbscript
' Text/Key Input
Const TEXT_REGISTRATION_DELAY = 50   ' ms for text to register
Const KEY_REGISTRATION_DELAY = 50    ' ms for key press to register

' Entry Processing
Const POST_ENTRY_WAIT = 100          ' ms after text entry
Const PRE_KEY_WAIT = 100             ' ms before special keys
Const POST_KEY_WAIT = 200            ' ms after special keys

' Prompt Detection
Const POLL_INTERVAL = 500            ' ms between prompt checks
Const PROMPT_DELAY = 200             ' ms after prompt detected

' Screen Operations
Const SCREEN_STABILITY_CHECK = 200   ' ms for simple operations
Const SCREEN_MODE_CHANGE = 500       ' ms for screen mode changes
Const SCREEN_TRANSITION = 500        ' ms for screen transitions
```

### Special Cases Requiring Longer Delays
```vbscript
' Screen Mode Changes
ADD_LABOR_OPERATION_DELAY = 800      ' Screen mode change
SOLD_HOURS_DELAY = 800               ' Complex form processing

' Stabilization Delays
ALL_LABOR_POSTED_PREWAIT = 1000      ' Major state transition
CHOOSE_OPTION_DELAY = 1000           ' Menu processing
```

## Rollback Strategy

If issues are encountered, use this incremental approach:

### Level 1: Increase Polling Only
```vbscript
Const POLL_INTERVAL = 750  ' Between 500 and 1000
```

### Level 2: Increase Key Delays
```vbscript
Const PRE_KEY_WAIT = 125   ' Between 100 and 150
Const POST_KEY_WAIT = 250  ' Between 200 and 350
```

### Level 3: Increase Specific Operations
```vbscript
' Only increase delays for specific failing operations
ADD_LABOR_OPERATION_DELAY = 1200  ' Increase by 50%
```

### Level 4: Full Rollback (Not Recommended)
Only if all incremental adjustments fail - consider investigating root cause instead.

## Monitoring Checklist

After deployment, monitor for:
- [ ] Timeout errors in logs
- [ ] Missed prompt detections
- [ ] Screen transition failures
- [ ] Data entry errors
- [ ] Unexpected automation stops

If any issues occur:
1. Identify the specific operation failing
2. Increase only that specific delay by 25-50%
3. Re-test
4. Document the adjustment and reason

## Success Metrics

Track these metrics to measure improvement:
- Average RO processing time
- Total daily throughput
- Error rate / retry rate
- Log entries for timeouts
- User satisfaction with speed

---
*This quick reference complements PERFORMANCE_OPTIMIZATION_SUMMARY.md*  
*For detailed analysis and implementation details, see the full summary*
