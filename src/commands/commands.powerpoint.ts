/*
 * Nasaq ribbon commands. Each button runs one of these functions directly
 * (ExecuteFunction), so there is no task pane.
 */

/* global Office PowerPoint console URL window */

import { alignToEdge, swapCenters } from "./geometry";
import type { Box, Edge } from "./geometry";
import { strings } from "./strings";
import type { CommandName } from "./strings";

const REQUIRED_API = { name: "PowerPointApi", version: "1.5" };

/** Error whose message is meant to be shown to the user as-is. */
class UserError extends Error {}

type SelectedShape = PowerPoint.Shape & Box;

/**
 * Loads the selected shapes on the current slide, in selection order
 * (the last item is the last shape the user selected).
 */
async function loadSelectedShapes(context: PowerPoint.RequestContext): Promise<SelectedShape[]> {
  const shapes = context.presentation.getSelectedShapes();
  shapes.load("items/id,items/left,items/top,items/width,items/height");
  await context.sync();
  return shapes.items;
}

function requireTwo(shapes: SelectedShape[], command: CommandName) {
  if (shapes.length < 2) {
    const t = strings();
    throw new UserError(t.needsTwo(t.commandNames[command], shapes.length));
  }
}

/**
 * Shows a message in a small dialog and resolves once the user closes it.
 * Function commands have no UI of their own, so this is how errors surface.
 */
function showMessage(message: string): Promise<void> {
  return new Promise((resolve) => {
    const url = new URL("message.html", window.location.href);
    url.searchParams.set("msg", message);
    Office.context.ui.displayDialogAsync(url.toString(), { height: 22, width: 28 }, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error(`Nasaq: ${message} (dialog failed: ${result.error.message})`);
        resolve();
        return;
      }
      const dialog = result.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, () => {
        dialog.close();
        resolve();
      });
      // Fires when the user closes the dialog with its own close button.
      dialog.addEventHandler(Office.EventType.DialogEventReceived, () => resolve());
    });
  });
}

/** Wraps a command body with the API check, error reporting and event.completed(). */
function command(body: (context: PowerPoint.RequestContext) => Promise<void>) {
  return async (event: Office.AddinCommands.Event) => {
    try {
      if (!Office.context.requirements.isSetSupported(REQUIRED_API.name, REQUIRED_API.version)) {
        throw new UserError(strings().needsApi(`${REQUIRED_API.name} ${REQUIRED_API.version}`));
      }
      await PowerPoint.run(body);
    } catch (error) {
      const message =
        error instanceof UserError
          ? error.message
          : strings().unexpected(String((error as Error).message ?? error));
      console.error(error);
      await showMessage(message);
    }
    event.completed();
  };
}

/** Resizes all selected shapes to the width and/or height of the last selected one. */
function match(name: CommandName, width: boolean, height: boolean) {
  return command(async (context) => {
    const shapes = await loadSelectedShapes(context);
    requireTwo(shapes, name);
    const target = shapes[shapes.length - 1];
    for (const shape of shapes.slice(0, -1)) {
      if (width) shape.width = target.width;
      if (height) shape.height = target.height;
    }
    await context.sync();
  });
}

const swap = command(async (context) => {
  const shapes = await loadSelectedShapes(context);
  if (shapes.length !== 2) {
    throw new UserError(strings().swapNeedsExactlyTwo(shapes.length));
  }
  const [a, b] = shapes;
  const [newA, newB] = swapCenters(a, b);
  a.left = newA.left;
  a.top = newA.top;
  b.left = newB.left;
  b.top = newB.top;
  await context.sync();
});

function align(edge: Edge) {
  return command(async (context) => {
    const shapes = await loadSelectedShapes(context);
    requireTwo(shapes, "align");
    // Align to the last selected shape, like Match Size.
    alignToEdge(shapes, edge, shapes[shapes.length - 1]).forEach((position, i) => {
      shapes[i].left = position.left;
      shapes[i].top = position.top;
    });
    await context.sync();
  });
}

/** Shape types that can hold text. */
const TEXT_SHAPE_TYPES = new Set<string>([
  "GeometricShape",
  "TextBox",
  "Placeholder",
  "Callout",
  "Freeform",
]);

/**
 * Sets paragraph alignment on the selected text, or on all text in the selected
 * shapes when no text is selected.
 */
function alignText(alignment: PowerPoint.ParagraphHorizontalAlignment) {
  return command(async (context) => {
    const range = context.presentation.getSelectedTextRangeOrNullObject();
    range.load("length");
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/type");
    await context.sync();

    if (!range.isNullObject && range.length > 0) {
      range.paragraphFormat.horizontalAlignment = alignment;
    } else {
      const textShapes = shapes.items.filter((shape) => TEXT_SHAPE_TYPES.has(shape.type));
      if (textShapes.length === 0) {
        throw new UserError(strings().needsText);
      }
      for (const shape of textShapes) {
        shape.textFrame.textRange.paragraphFormat.horizontalAlignment = alignment;
      }
    }
    await context.sync();
  });
}

// Register the add-in commands with the Office host application.
// The names must match the <FunctionName> values in manifest.xml.
Office.onReady(() => {
  Office.actions.associate("matchSize", match("matchSize", true, true));
  Office.actions.associate("matchWidth", match("matchWidth", true, false));
  Office.actions.associate("matchHeight", match("matchHeight", false, true));
  Office.actions.associate("swap", swap);
  Office.actions.associate("alignLeft", align("left"));
  Office.actions.associate("alignRight", align("right"));
  Office.actions.associate("alignTop", align("top"));
  Office.actions.associate("alignBottom", align("bottom"));
  Office.actions.associate("alignCenter", align("center"));
  Office.actions.associate("alignMiddle", align("middle"));
  Office.actions.associate(
    "textAlignLeft",
    alignText(PowerPoint.ParagraphHorizontalAlignment.left)
  );
  Office.actions.associate(
    "textAlignCenter",
    alignText(PowerPoint.ParagraphHorizontalAlignment.center)
  );
  Office.actions.associate(
    "textAlignRight",
    alignText(PowerPoint.ParagraphHorizontalAlignment.right)
  );
  Office.actions.associate(
    "textAlignJustify",
    alignText(PowerPoint.ParagraphHorizontalAlignment.justify)
  );
});
