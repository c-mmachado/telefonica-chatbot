import { TurnContext } from "botbuilder";
import { DialogTurnResult } from "botbuilder-dialogs";

import { RunnableDialog } from "./dialog";

export enum OAuthStatus {
  Waiting,
  Complete,
  Failed,
}

export interface DialogManager {
  registerDialog(dialog: RunnableDialog): void;

  runDialog(
    context: TurnContext,
    dialogName: string,
    data?: any
  ): Promise<DialogTurnResult>;

  continueDialog(
    context: TurnContext,
    dialogName: string
  ): Promise<DialogTurnResult>;

  stopDialog(
    context: TurnContext,
    dialogName: string
  ): Promise<DialogTurnResult>;

  findDialog(dialogName: string): RunnableDialog | null;
}

export class DefaultDialogManager implements DialogManager {
  private readonly _dialogs: RunnableDialog[] = [];

  constructor() {}

  public registerDialog(dialog: RunnableDialog): void {
    if (this.findDialog(dialog.Name)) {
      throw new Error(`Dialog '${dialog.Name}' is already registered`);
    }
    this._dialogs.push(dialog);
  }

  public async runDialog(
    context: TurnContext,
    dialogName: string,
    data?: {
      sequenceId?: string;
    } & any
  ): Promise<DialogTurnResult> {
    console.debug(`Attempting to run dialog with name '${dialogName}'`);

    const dialog = this.findDialog(dialogName);
    if (!dialog) {
      throw new Error(`Unable to retrieve dialog with name: '${dialogName}'`);
    }
    return await dialog.run(context, data);
  }

  public async continueDialog(
    context: TurnContext,
    dialogName: string
  ): Promise<DialogTurnResult> {
    console.debug(`Attempting to continue dialog with name '${dialogName}'`);

    const dialog = this.findDialog(dialogName);
    if (!dialog) {
      throw new Error(`Unable to retrieve dialog with name: '${dialogName}'`);
    }
    return await dialog.continue(context);
  }

  public async stopDialog(
    context: TurnContext,
    dialogName: string
  ): Promise<DialogTurnResult> {
    console.debug(`Attempting to stop dialog with name '${dialogName}'`);

    const dialog = this.findDialog(dialogName);
    if (!dialog) {
      throw new Error(`Unable to retrieve dialog with name: '${dialogName}'`);
    }
    return await dialog.stop(context);
  }

  public findDialog(dialogName: string): RunnableDialog | null {
    const dialog = this._dialogs.find(
      (dialog: RunnableDialog) => dialog.Name === dialogName
    );
    return dialog || null;
  }
}
