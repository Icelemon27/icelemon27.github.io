// 示例：验证公式是否只包含加减乘除和数字、括号、空格等简单字符
// 这里只是一个示例，实际应用中可能需要更复杂的正则和逻辑
function validateFormula(formula) {
    // 允许的字符：数字、运算符、括号、空格、小数点（.）、字母（用于单元格引用，假设仅限于单个字母和数字，如 A1）
    const allowedRegex = /^[A-Za-z0-9+\-*/().\s=]+$/;
    return allowedRegex.test(formula);
  }
  
  document.getElementById('validateFormulaBtn').addEventListener('click', function() {
    // 示例公式，可替换为实际从单元格中获取的公式
    const sampleFormula = "Price * Quantity"; // 假设这是用户输入的公式
    if (!validateFormula(sampleFormula)) {
      document.getElementById('formulaWarning').style.display = "block";
    } else {
      document.getElementById('formulaWarning').style.display = "none";
      alert("公式验证通过！");
    }
  });
  