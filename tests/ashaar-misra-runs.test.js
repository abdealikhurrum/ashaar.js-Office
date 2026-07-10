const assert = require("assert");
const Ashaar = require("../src/vendor/ashaar");

// Empty/whitespace child nodes are dropped; order and styling preserved.
{
  const specs = Ashaar.misraRunSpecs([
    { text: "درد ",  fontKey: "Amiri/400/normal", fontSize: 16 },
    { text: "دل",    fontKey: "Amiri/700/normal", fontSize: 16 },
    { text: "   ",   fontKey: "Amiri/400/normal", fontSize: 16 },
  ]);
  assert.equal(specs.length, 2);
  assert.equal(specs[0].fontKey, "Amiri/400/normal");
  assert.equal(specs[1].text, "دل");
  assert.equal(specs[1].fontSize, 16);
}

// Empty input → [].
assert.deepEqual(Ashaar.misraRunSpecs([]), []);
assert.deepEqual(Ashaar.misraRunSpecs(), []);

console.log("ashaar-misra-runs tests passed");
