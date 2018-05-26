export interface TemplatePipe {
  pipeName: string;
  pipeParameters: string[];
}

export class TemplateExpression {
  public rawExpression: string;
  public expression: string;
  public valueName: string;
  public pipes: TemplatePipe[];

  constructor(rawExpression: string, expression: string) {
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
