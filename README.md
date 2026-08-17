# Minecraft server in Word

A working Minecraft 1.8.9 server implemented entirely in Microsoft Word VBA. Real client with a real protocol and no server software. The process listening on port 25565 is WINWORD.EXE.

The world is drawn on the page as an isometric render made of cube autoshapes, so it moves as you walk. The part that's actually Word is the margin. By the end the document is a full edit history of a Minecraft session.

## Why

Eighth one. Excel, Outlook, OBS, Obsidian, Blender, PowerPoint, VLC, now Word.

Word's job is recording changes to a document. This uses it to record changes to a world.

Nobody has made Word speak the Minecraft protocol as far as I'm aware.

## What it does

- Server list ping with MOTD, player count and a favicon
- Offline-mode login, no encryption
- Flat world, 5x5 chunks of bedrock, dirt and grass
- Creative mode, so you can fly and place things
- Keep-alives
- A live isometric render floating on the page, 140 cube autoshapes, redrawn as you move
- Blocks you place appear as cubes, coloured by block ID, and vanish when you break them
- Every block placed or broken written into the document as a tracked revision, named and with coordinates
- A Heading 2 dropped for each new chunk you enter, so the navigation pane populates
- Sub-block scrolling, so walking slides the world rather than snapping a tile at a time
- A live status line showing player, position, facing, block count and packet counts

## What it doesn't do

Almost everything else. No block updates on the server side, no entities, no mobs, no real physics, no chat commands, no second player. Blocks you place exist only in your client.

Three limits specific to this build:

**The render is a window, not the world.** It shows a 9x9 area around you, not all 25 chunks. Anything further out isn't drawn.

**Breaking terrain isn't reflected.** The floor is drawn fresh every repaint from a fixed height, so only blocks you placed can be removed.

**The log is capped at 30 lines.** Word with hundreds of tracked revisions gets very slow, so the oldest lines are deleted as new ones arrive.

It's a server in the sense that a client connects to it and receives a world. Set your expectations accordingly.

## Requirements

Windows, and Word with VBA. Not Word for Mac, where VBA runs inside the macOS App Sandbox and the Winsock calls compile and then fail at runtime. My video examples run on Windows 11 ARM in UTM.

Minecraft Java 1.8.9, protocol 47.

## Setup

**1. New document, saved as `.docm`.** Macro-Enabled Document, not `.docx`. A `.docx` throws the code away when you close it.

**2. Import the module.** Alt+F11, then File > Import File, and pick `WordMCServer.bas`.

**3. Enable macros.** File > Options > Trust Center > Trust Center Settings > Macro Settings.

**4. Run `Setup` once.** Ctrl+G in the VBA editor, type `Setup`, Enter. That builds the page, the 140 shape pool, the status line and the build log.

**5. Set the view up before you start.** This is the difference between the good version and a thin grey line nobody notices:

- View tab, **Print Layout**
- Review tab, markup dropdown set to **All Markup**
- Review tab, Show Markup > Balloons > **Show Revisions in Balloons**

**6. Run `StartServer`.** Ctrl+G again, or put the macros on the Quick Access Toolbar via File > Options > Quick Access Toolbar with **Choose commands from** set to Macros. Save them to the document rather than to Word and they travel with the file.

**7. Connect** to `localhost:25565` in Minecraft 1.8.9 via Direct Connect. Allow the Windows firewall prompt on Private networks.

Optional: a 64x64 PNG named `favicon.png` beside the saved `.docm` shows as the server icon in the multiplayer list. Must be exactly 64x64 or the client drops it. Save the document before starting or there's no path to look in.

Re-running `Setup` is safe. It looks for the shapes and bookmarks by name and rebinds to them rather than rebuilding.

## How it works

VBA has no networking. It has `Declare`, and `Declare` reaches `ws2_32.dll`, which is Winsock, which is what everything on Windows goes through to reach the network.

Word has no `Application.OnTime`, so the server runs on a Windows `SetTimer` callback with `AddressOf`, the same as the Outlook and PowerPoint builds. Every 20ms it does one pass of accept, read, flush and keep-alive, then returns.

**VarInts.** The protocol's variable width integer format, used for every packet length and ID. VBA has no unsigned right shift, so `URShift7` fakes one.

**Packet framing.** Length, then packet ID, then payload. TCP is a stream and not a sequence of messages, so packets arrive split across reads or several at once.

**Chunks.** The reason this targets 1.8.9. In 1.8 a chunk section is a flat array of `(id << 4) | meta` shorts, then block light, then sky light. 12,544 bytes for one section plus biome data, sent uncompressed. Modern versions use palette-encoded, bit packed longs and expect zlib, and VBA has no zlib. The join sequence is 25 chunks, so 313,600 bytes in one burst.

**Rendering.** Cube shapes are created once and then only moved and recoloured. Each repaint builds a draw list from the floor around you plus any placed blocks in range, sorts it back to front with the key `(x + z) * 128 + y`, then assigns it to the pool in order. Because the pool was created in that order, z-order comes out correct for free. Repaints only happen when you cross a block boundary or a block changes.

Word shapes anchor to a paragraph by default, which would drag the render around whenever the log grows. `RelativeHorizontalPosition` and `RelativeVerticalPosition` are set to the page, and `WrapFormat.Type` to `wdWrapFront`, so the coordinates are absolute and the cubes float over the text.

**Position decode.** Player Digging and Player Block Placement both carry a packed 64-bit Position, 26 bits of X, 12 of Y, 26 of Z. VBA's Long is 32-bit, so it's unpacked byte by byte and sign-extended by hand.

| Packet | What it carries | What the document does with it |
| --- | --- | --- |
| `0x04` Player Position | position | scrolls the render |
| `0x05` Player Look | yaw, pitch | updates the status line |
| `0x06` Player Position And Look | both | both |
| `0x07` Player Digging | status, block position | removes a cube, logs a tracked deletion |
| `0x08` Player Block Placement | clicked position, face, held item | adds a cube, logs a tracked insertion |

## Notes from building eight of these

The protocol half was a straight lift from PowerPoint. Eight builds in, that part is solved. What each host still costs you is the render layer, and Word's is the fiddliest so far because the display surface and the document are the same object.

Three things had to be designed around Track Changes.

**Shape creation has to happen with tracking off.** Otherwise 140 shape insertions become 140 revisions before a client has even connected, and the margin is full before the video starts. `Setup` builds the pool, `StartServer` turns tracking on afterwards.

**The status line updates untracked.** It rewrites several times a second, and tracked it would bury every real event under a wall of noise. Only block placements, breaks and chunk headings get recorded, which is the whole reason the log reads as a build history rather than a diff.

**Trimming the log has to disable tracking first.** Deleting a tracked insertion while tracking is on doesn't remove it, it marks it deleted, and the document grows forever while getting slower. The trim toggles tracking off, deletes and puts it back.

Word also reflows the page on every shape change, so it's slower than PowerPoint at the same pool size. 140 shapes over a 9x9 area is where it settled. If it drags on your machine, `RADIUS 3` and `POOL 90` is comfortable.

## Previous episodes

- Excel: [github.com/hostdit/minecraft-server-in-excel](https://github.com/hostdit/minecraft-server-in-excel)
- Outlook: [github.com/hostdit/minecraft-server-in-outlook](https://github.com/hostdit/minecraft-server-in-outlook)
- OBS: [github.com/hostdit/minecraft-server-in-obs](https://github.com/hostdit/minecraft-server-in-obs)
- Obsidian: [github.com/hostdit/minecraft-server-in-obsidian](https://github.com/hostdit/minecraft-server-in-obsidian)
- Blender: [github.com/hostdit/minecraft-server-in-blender](https://github.com/hostdit/minecraft-server-in-blender)
- PowerPoint: [github.com/hostdit/minecraft-server-in-powerpoint](https://github.com/hostdit/minecraft-server-in-powerpoint)
- VLC: [github.com/hostdit/minecraft-server-in-vlc](https://github.com/hostdit/minecraft-server-in-vlc)

## Licence

MIT. Do what you like with it.
