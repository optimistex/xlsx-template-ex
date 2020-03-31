export interface TemplatePipe {
  pipeName: string;
  pipeParameters: string[];
}

export class TemplateExpression {
  public valueName: string;
  public pipes: TemplatePipe[] = [];

  constructor(public rawExpression: string, expression: string) {
    // this.rawExpression = rawExpression;
    const expressionParts = expression.split("|").map(e => e.trim());
    this.valueName = expressionParts[0];
    const pipes = expressionParts.slice(1);
    pipes.forEach((pipe: string) => {
      const pipeParts = pipe.split(":").map(p => p.trim());
      this.pipes.push({ pipeName: pipeParts[0], pipeParameters: pipeParts.slice(1) });
    });
  }
}
