# PostFinalCharges Test Suite Documentation

This directory contains comprehensive tests for the PostFinalCharges automation, with special focus on **Default Value Detection** and **Real-World Prompt Variations**.

## 🚨 Critical Testing Guidelines for AI Agents

### Test-Driven Prevention Strategy

When working on prompt detection, default value logic, or screen interactions, **always follow these patterns to prevent e2e test failures**:

#### 1. Real-World Prompt Coverage Requirements
**REQUIRED**: Include these prompt variation types in ALL new tests:

```vbs
' ✅ REQUIRED TEST PATTERNS:
' Simple format (basic functionality)
"TECHNICIAN (72925)?" → Should detect default "72925"

' Intervening text (CRITICAL - this pattern was missed and caused production bugs)  
"OPERATION CODE FOR LINE A, L1 (I)?" → Should detect default "I"

' Multiple context elements
"TECHNICIAN FOR JOB 12345, LINE A (T999)?" → Should detect default "T999"

' Multiple parentheses (edge case)
"COST CENTER (MAIN) FOR LINE A (CC123)?" → Should detect default "CC123"

' Negative cases (no defaults)
"TECHNICIAN?" → Should NOT detect any default
"TECHNICIAN ()?" → Should NOT accept empty defaults
```

#### 2. Mandatory Test Case Categories
Every prompt processing change MUST include tests for:

- **Simple patterns** (verify existing functionality works)
- **Intervening text patterns** (prevent the "OPERATION CODE FOR LINE A, L1" class of bugs)
- **Edge cases** (empty values, malformed prompts, multiple parentheses)
- **Real-world variations** (based on actual screen content from e2e testing)

#### 3. Test Documentation Standards
Each test case MUST document its purpose and real-world source:

```vbs
' Test 11: OPERATION CODE FOR LINE with intervening text (THE ACTUAL BUG!)
' Source: Found during e2e testing, missed by unit tests
' Pattern: keyword + descriptive text + (default_value)?
' This test prevents future regressions of intervening text scenarios
```

## Test Files

### Enhanced Default Value Tests
- **`test_default_value_detection.vbs`** - **15 comprehensive test cases** including intervening text patterns that would have caught the production bug
- **`test_operation_code_default.vbs`** - Specific tests for the "OPERATION CODE FOR LINE" issue
- **`test_bug_prevention.vbs`** - Demonstrates how enhanced tests catch regex pattern failures

### Integration & Mock Tests
- **`test_default_value_detection.vbs`** - Unit tests for the `HasDefaultValueInPrompt()` function
- **`test_default_value_integration.vbs`** - Integration tests using MockBzhao to simulate the complete flow
- **`run_default_value_tests.vbs`** - Test runner specifically for the default value tests

### Existing Tests
- **`test_mock_bzhao.vbs`** - Tests the MockBzhao mock framework
- **`test_integration.vbs`** - Integration tests for the main script
- **`test_prompt_detection.vbs`** - Tests for prompt detection timing

## Running the Tests

### Option 1: Run All Tests
```cmd
cd utilities\tests
cscript.exe run_all_tests.vbs
```

### Option 2: Run Only Default Value Tests
```cmd
cd utilities\tests
cscript.exe run_default_value_tests.vbs
```

### Option 3: Run Individual Tests
```cmd
cd utilities\tests
cscript.exe test_default_value_detection.vbs
cscript.exe test_default_value_integration.vbs
```

## What the Bugfix Tests Verify

### Unit Tests (`test_default_value_detection.vbs`)
- ✅ `TECHNICIAN (72925)?` → Detects default value "72925"
- ✅ `ACTUAL HOURS (117)?` → Detects default value "117" 
- ✅ `SOLD HOURS (0)?` → Detects default value "0"
- ✅ `TECHNICIAN?` → No default detected (correctly)
- ✅ `TECHNICIAN ()?` → No default detected (correctly)

### Integration Tests (`test_default_value_integration.vbs`)  
- ✅ Simulates complete prompt sequence with defaults
- ✅ Verifies that only ENTER is sent (no hardcoded values)
- ✅ Tests the behavior comparison (old vs new)

## The Bug That Was Fixed

**Before the fix:**
- `TECHNICIAN (72925)?` → Script sends "99" + Enter → Result: Technician 99 ❌
- `ACTUAL HOURS (117)?` → Script sends "0" + Enter → Result: 0 hours ❌

**After the fix:**  
- `TECHNICIAN (72925)?` → Script sends Enter only → Result: Technician 72925 ✅
- `ACTUAL HOURS (117)?` → Script sends Enter only → Result: 117 hours ✅

## Test Results

All tests should pass if the bugfix is working correctly. The tests verify:

1. **Function Logic**: The `HasDefaultValueInPrompt()` function correctly identifies when a prompt contains a default value
2. **Prompt Configuration**: The prompt dictionary correctly marks relevant prompts with `AcceptDefault = True`
3. **Integration**: The complete prompt processing flow accepts defaults instead of overwriting them

## Future Test Additions

Consider adding tests for:
- Edge cases with unusual default value formats
- Performance testing with large screen buffers
- Real BlueZone integration tests (require actual BlueZone setup)
- Regression tests for related prompt patterns