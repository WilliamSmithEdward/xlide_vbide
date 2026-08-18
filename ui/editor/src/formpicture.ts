/**
 * A form's PICTURES, drawn: MSForms' placement properties turned into CSS, and nothing else.
 *
 * A module of its own because it is the one part of the canvas that holds no view state. It
 * takes a picture from the projection and an element to put it on, and it knows two things: the
 * two families MSForms has, and how each maps onto a background or a flex line. Nothing here
 * reads the open page, the zoom, the selection or the document - which is what made it separable
 * from designerview.ts once the canvas had grown past three thousand lines (2026-08-16).
 *
 * The mapping is exact rather than approximate, which is the point worth defending: a SURFACE
 * picture's three properties are a background's three properties, and a CAPTION picture's twelve
 * positions are four flex directions by three cross alignments. Neither needed a layout engine
 * and neither is a guess.
 */
import type { FormMarkupPicture } from "./bridge.js";

/** The kinds whose picture is their SURFACE rather than something beside a caption. Each paints
 * on a different part of itself - a Frame on its client, a Page on its body - so each does it
 * where it is drawn, and this set is what keeps the caption painter off them. */
export const SURFACE_PICTURE_TYPES = new Set(["Image", "Frame", "Page", "MultiPage", "TabStrip"]);

/** fmPictureAlignment, as a background-position: 0 top-left, 1 top-right, 2 centre, 3
 * bottom-left, 4 bottom-right. Centre is the default and the fallback. */
const PICTURE_ALIGNMENTS = ["left top", "right top", "center center", "left bottom", "right bottom"];

/** One background layer. A layer rather than four style writes because the FORM's picture shares
 * its element with the grid, and two things on one background have to be composed. */
export interface PictureLayer {
  image: string;
  size: string;
  position: string;
  repeat: string;
}

/**
 * A picture as a background layer: a form's, an Image's, a Frame's client, a Page's body.
 *
 * The three placement properties map onto CSS exactly, which is the reason this is four lines
 * rather than a layout engine. fmPictureSizeMode: 0 clips (natural size, the element's own
 * overflow does the clipping), 1 stretches to the box, 3 zooms to fit inside it. Tiling repeats
 * from the alignment's own corner, which is what a repeating background does.
 */
export function pictureLayer(picture: FormMarkupPicture): PictureLayer {
  return {
    image: `url("${picture.src}")`,
    repeat: picture.tiling ? "repeat" : "no-repeat",
    position: PICTURE_ALIGNMENTS[picture.alignment ?? 2] ?? "center center",
    size: picture.sizeMode === 1 ? "100% 100%" : picture.sizeMode === 3 ? "contain" : "auto",
  };
}

/** The layer painted straight onto an element that has nothing else on its background. */
export function paintPictureSurface(element: HTMLElement, picture: FormMarkupPicture): void {
  const layer = pictureLayer(picture);
  element.style.backgroundImage = layer.image;
  element.style.backgroundRepeat = layer.repeat;
  element.style.backgroundPosition = layer.position;
  element.style.backgroundSize = layer.size;
}

/**
 * A picture that sits WITH a caption, where fmPicturePosition says.
 *
 * All twelve positions, exactly, because they are four directions by three cross alignments and
 * flexbox is four directions by three cross alignments: 0-2 left, 3-5 right, 6-8 above, 9-11
 * below, each triple aligned start, centre, end along the other axis. 12 is the odd one - the
 * picture BEHIND the caption - and it is a background rather than a child.
 *
 * An `<img>` rather than a background for the eleven, because a picture beside a caption is
 * drawn at its natural size and an `img` is the one element that knows what that is. The canvas
 * scales a point to 4/3 of a pixel and a picture's pixel IS a CSS pixel, which is also what the
 * runtime does with it at 96dpi.
 */
export function dressWithPicture(box: HTMLElement, picture: FormMarkupPicture): void {
  const position = picture.position ?? 7;
  if (position === 12) {
    box.style.backgroundImage = `url("${picture.src}")`;
    box.style.backgroundRepeat = "no-repeat";
    box.style.backgroundPosition = "center center";
    return;
  }

  const image = document.createElement("img");
  image.className = "dc-picture";
  image.src = picture.src;
  image.alt = "";
  image.draggable = false;

  /*
   * A CAPTION PICTURE IS DRAWN THROUGH A COLOUR KEY, and a surface picture is not. Measured
   * 2026-08-17 off the running form, both controls wearing the same black-backed logo:
   *
   *   the OK button (PicturePosition)  the runtime keys the black out and the button face shows
   *                                    through, with a speckled halo where the anti-aliased
   *                                    near-black pixels missed the key
   *   the Badge (PictureSizeMode)      the runtime draws the black field solid, letterboxed
   *
   * So this is not "black is transparent" and it is not something the loader does to the picture:
   * the two surfaces are handed the SAME pixels and MSForms draws them two different ways. The
   * canvas drew both opaque, which is right for the Badge and is the disparity the owner reported
   * five times over that day about the button.
   *
   * EXACTLY-MATCHING PIXELS ONLY, which is not a shortcut - it is what produces the halo. A
   * tolerance would clean up the speckles and stop matching the thing being matched.
   */
  const keyOutBackground = (): void => {
    if (image.naturalWidth === 0 || image.dataset.keyed === "yes") {
      return;
    }

    const sheet = document.createElement("canvas");
    sheet.width = image.naturalWidth;
    sheet.height = image.naturalHeight;
    const ink = sheet.getContext("2d", { willReadFrequently: false });
    if (ink === null) {
      return;
    }

    ink.drawImage(image, 0, 0);

    let field: ImageData;
    try {
      field = ink.getImageData(0, 0, sheet.width, sheet.height);
    } catch {
      // A picture the page may not read back. Nothing to key, and the opaque draw still stands.
      return;
    }

    // The top-left pixel is the key, which is the corner MSForms takes it from.
    const pixels = field.data;
    const [red, green, blue] = [pixels[0], pixels[1], pixels[2]];
    for (let at = 0; at < pixels.length; at += 4) {
      if (pixels[at] === red && pixels[at + 1] === green && pixels[at + 2] === blue) {
        pixels[at + 3] = 0;
      }
    }

    ink.putImageData(field, 0, 0);
    image.dataset.keyed = "yes";
    image.src = sheet.toDataURL("image/png");
  };

  /*
   * A PICTURE TOO BIG FOR ITS CONTROL IS STRETCHED, not letterboxed and not clipped, because
   * that is what MSForms does - measured off the running form twice (the owner: "see distortion
   * of image over button in live form", then "native form is stretch, and xlide canvas is
   * truncating"). The fixture's OK button is 72x24 points wearing a 256x256 icon, and the
   * runtime squashes the whole logo into it.
   *
   * It cannot be said in CSS alone: `object-fit` needs a box, and the box is only over-large
   * once the intrinsic size is known - which the payload does not carry and the IMG does. So the
   * decision waits for the load and then only fires when the picture really does not fit. A
   * picture that fits keeps its natural size, which is what PicturePosition is for.
   */
  const stretchIfOversized = (): void => {
    const room = box.getBoundingClientRect();
    if (room.width <= 0 || room.height <= 0 || image.naturalWidth === 0) {
      return;
    }

    if (image.naturalWidth > room.width || image.naturalHeight > room.height) {
      // It COVERS the control rather than sharing a flex line with the caption. Stretching it in
      // place pushed the caption out the side and clipped it, which was a third wrong answer
      // (the owner: "still a disparity in lower button"); the runtime paints the picture over the
      // whole face and the caption goes under it, so the canvas takes it out of the flow.
      box.style.position = "relative";
      image.style.position = "absolute";
      image.style.inset = "0";
      image.style.width = "100%";
      image.style.height = "100%";
      image.style.objectFit = "fill";
    }
  };

  // The key first, because it replaces the src and so raises load again; the second pass sees
  // `keyed` and only measures. Both steps need the natural size, which is why both wait for load.
  const dress = (): void => {
    keyOutBackground();
    stretchIfOversized();
  };

  image.addEventListener("load", dress);
  if (image.complete) {
    dress();
  }

  box.classList.add("dc-with-picture");
  box.style.flexDirection = position <= 5 ? "row" : "column";
  box.style.alignItems = ["flex-start", "center", "flex-end"][position % 3] ?? "center";

  // Left and above come FIRST, right and below last - which is also the order the two triples
  // sit in the enumeration, so the placement reads off the number.
  const before = position <= 2 || (position >= 6 && position <= 8);
  if (before) {
    box.prepend(image);
  } else {
    box.appendChild(image);
  }
}
