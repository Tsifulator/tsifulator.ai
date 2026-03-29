"""
Generate olive wreath SVG matching reference image.
2 branches in a swirling pattern, each curving ~270° around a circle.
Long pointed olive leaves growing from stems. Small berries.
"""
import math

W, H = 512, 512
CX, CY = 256, 256
R = 120  # radius of the circular path the stems follow

def point_on_circle(cx, cy, r, angle_deg):
    a = math.radians(angle_deg)
    return cx + r * math.cos(a), cy + r * math.sin(a)

def leaf_path(bx, by, angle, length=38, width=11):
    """Create a pointed olive leaf as an SVG path."""
    a = math.radians(angle)
    # Tip of leaf
    tx = bx + length * math.cos(a)
    ty = by + length * math.sin(a)
    # Control points for the two sides (perpendicular offset)
    perp = a + math.pi / 2
    # Midpoint
    mx = bx + length * 0.45 * math.cos(a)
    my = by + length * 0.45 * math.sin(a)
    # Bulge points
    c1x = mx + width * math.cos(perp)
    c1y = my + width * math.sin(perp)
    c2x = mx - width * math.cos(perp)
    c2y = my - width * math.sin(perp)
    return f"M{bx:.1f},{by:.1f} Q{c1x:.1f},{c1y:.1f} {tx:.1f},{ty:.1f} Q{c2x:.1f},{c2y:.1f} {bx:.1f},{by:.1f}Z"

def generate_branch(start_angle, end_angle, direction=1):
    """Generate leaves and berries along a curved branch stem."""
    elements = []

    # Number of leaf pairs
    n_leaves = 12
    angles = []
    for i in range(n_leaves):
        t = i / (n_leaves - 1)
        angle = start_angle + (end_angle - start_angle) * t
        angles.append(angle)

    # Stem path - series of points along the arc
    stem_points = []
    n_stem = 40
    for i in range(n_stem + 1):
        t = i / n_stem
        angle = start_angle + (end_angle - start_angle) * t
        # Slight spiral - radius decreases slightly toward the end
        r = R + 5 * math.sin(t * math.pi)
        x, y = point_on_circle(CX, CY, r, angle)
        stem_points.append((x, y))

    # Build stem as smooth path
    stem_d = f"M{stem_points[0][0]:.1f},{stem_points[0][1]:.1f}"
    # Use cubic bezier through groups of points
    for i in range(1, len(stem_points)):
        stem_d += f" L{stem_points[i][0]:.1f},{stem_points[i][1]:.1f}"
    elements.append(f'<path d="{stem_d}" stroke="#0D5EAF" stroke-width="2.8" fill="none" stroke-linecap="round"/>')

    # Generate leaves at each position
    for i, angle_deg in enumerate(angles):
        t = i / (n_leaves - 1)
        # Position on the arc
        r_stem = R + 5 * math.sin(t * math.pi)
        bx, by = point_on_circle(CX, CY, r_stem, angle_deg)

        # Tangent direction along the arc
        tangent = angle_deg + 90 * direction

        # Leaf size varies slightly
        size_factor = 0.75 + 0.5 * math.sin(t * math.pi)  # bigger in middle
        leaf_len = 34 * size_factor
        leaf_w = 9 * size_factor

        # Outer leaf - points outward from center
        outward_angle = angle_deg  # radial direction (away from center)
        # Tilt the leaf along the tangent
        outer_angle = outward_angle + direction * 25
        outer_path = leaf_path(bx, by, outer_angle, leaf_len, leaf_w)
        elements.append(f'<path d="{outer_path}"/>')

        # Inner leaf - points inward-ish, along tangent
        inner_angle = outward_angle + 180 - direction * 35
        inner_len = leaf_len * 0.85
        inner_w = leaf_w * 0.85
        inner_path = leaf_path(bx, by, inner_angle, inner_len, inner_w)
        elements.append(f'<path d="{inner_path}"/>')

        # Extra side leaves for density (every other position)
        if i % 2 == 0:
            extra_angle = outward_angle + direction * 60
            extra_path = leaf_path(bx, by, extra_angle, leaf_len * 0.7, leaf_w * 0.7)
            elements.append(f'<path d="{extra_path}"/>')

        if i % 3 == 0:
            extra_angle2 = outward_angle + 180 + direction * 15
            extra_path2 = leaf_path(bx, by, extra_angle2, leaf_len * 0.65, leaf_w * 0.65)
            elements.append(f'<path d="{extra_path2}"/>')

    # Berries - every 3rd position, slightly offset
    for i in range(1, n_leaves - 1, 3):
        t = i / (n_leaves - 1)
        angle_deg = angles[i]
        r_stem = R + 5 * math.sin(t * math.pi)
        bx, by = point_on_circle(CX, CY, r_stem, angle_deg)
        # Offset berry slightly along tangent
        offset_angle = math.radians(angle_deg + 90 * direction)
        ox = bx + 8 * math.cos(offset_angle)
        oy = by + 8 * math.sin(offset_angle)
        elements.append(f'<circle cx="{ox:.1f}" cy="{oy:.1f}" r="4"/>')

    return elements

# Build SVG
svg_parts = [
    f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {W} {H}" fill="#0D5EAF">',
]

# Branch 1: sweeps from bottom-left (~210°) counterclockwise to upper-right (~330°)
# Actually, looking at the reference more carefully:
# Branch 1 starts around 7 o'clock (210°), sweeps clockwise up to about 1 o'clock (30°)
# Branch 2 starts around 1 o'clock (30°), sweeps clockwise down to about 7 o'clock (210°)
# Together they create a full circle with a swirl

branch1 = generate_branch(210, -30, direction=1)   # ~240° arc clockwise (going up-right)
branch2 = generate_branch(30, 150 + 60, direction=-1)  # ~240° arc clockwise (going down-left)

svg_parts.extend(branch1)
svg_parts.extend(branch2)
svg_parts.append('</svg>')

svg = '\n'.join(svg_parts)

with open('/Users/nicholastsiflikiotis/tsifulator.ai/brand/tsifl-wreath-icon.svg', 'w') as f:
    f.write(svg)

print("Generated wreath SVG")
print(f"Branches: 2, Leaf pairs per branch: ~12")
