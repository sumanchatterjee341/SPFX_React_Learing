export class Validation {
  public TextFieldValidation(input: string): string {
    if (input.length == 0) {
      return "Field cannot be empty.";
    }
    if (input.length > 256) {
      return "Field Value length is too long.";
    }
    return "";
  }
}
