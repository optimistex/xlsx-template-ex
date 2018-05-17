/**
 * @property {string} rawExpression
 * @property {string} expression
 * @property {string} valueName
 * @property {Array<{pipeName: string, pipeParameters: string[]}>} pipes
 */
class TemplateExpression {
    /**
     * @param {string} rawExpression
     * @param {string} expression
     */
    constructor(rawExpression, expression) {
        this.rawExpression = rawExpression;
        this.expression = expression;
        const expressionParts = this.expression.split('|');
        this.valueName = expressionParts[0];
        this.pipes = [];
        const pipes = expressionParts.slice(1);
        pipes.forEach(pipe => {
            const pipeParts = pipe.split(':');
            this.pipes.push({pipeName: pipeParts[0], pipeParameters: pipeParts.slice(1)});
        });
    }
}

module.exports = TemplateExpression;