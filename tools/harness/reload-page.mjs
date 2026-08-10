/*
 * Puts the bundle now on disk in front of every live editor.
 *
 * The shim serves the page over loopback from a folder, reading each file as it is asked for, so
 * a page change needs no republish and no restart - only a reload. Restarting the host is for
 * SHIM changes, because a host holds an add-in library open for its lifetime.
 *
 * Reloading through the api rather than by hand because it WAITS for the page to say it is ready:
 * a reload followed by a guessed sleep reports on the page that is going away.
 *
 *   node reload-page.mjs
 */

import { discover } from "./xlide-api.mjs";

const instances = await discover();

if (instances.length === 0) {
  console.log("no live editor to reload; the files are in place for the next start");
  process.exit(0);
}

let failed = 0;

for (const instance of instances) {
  try {
    const answer = await instance.api.reload();
    const stamp = answer.stamp ?? answer.pageBuildStamp ?? JSON.stringify(answer);
    console.log(`pid ${instance.pid}: now running page ${stamp}`);
  } catch (error) {
    failed++;
    console.log(`pid ${instance.pid}: reload failed - ${error.message}`);
  }
}

process.exit(failed > 0 ? 1 : 0);
