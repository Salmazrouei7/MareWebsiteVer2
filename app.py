import json
import math
import re
from collections import defaultdict

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components


st.set_page_config(
    page_title="Arabian Mare Dam-Line Family Tree",
    layout="wide",
)

FILE_NAME = "Horses.xlsx"
DAM_SHEET = "Dam Lines"

st.title("ARABIAN MARE DAM-LINE FAMILY TREE")
st.caption("Filter race results and show only matching maternal branches")


# ============================================================
# READ EXCEL DATA
# ============================================================
try:
    dam_df = pd.read_excel(FILE_NAME, sheet_name=DAM_SHEET)
    excel_file = pd.ExcelFile(FILE_NAME)
except Exception as exc:
    st.error(f"Could not open {FILE_NAME}: {exc}")
    st.stop()


race_sheet = None

for sheet_name in excel_file.sheet_names:
    lower_name = sheet_name.lower()

    if "race" in lower_name and "result" in lower_name:
        race_sheet = sheet_name
        break


if race_sheet is None:
    st.error("Could not find the Race Results sheet.")
    st.stop()


race_df = pd.read_excel(FILE_NAME, sheet_name=race_sheet)


# ============================================================
# DAM-LINE SHEET
# ============================================================
# Columns A:J contain:
# Mare, Dam, earlier dams ..., Family/Foundation mare

if dam_df.shape[1] < 10:
    st.error("The Dam Lines sheet must contain at least 10 columns.")
    st.stop()


chain_cols = list(dam_df.columns[:10])
family_col = dam_df.columns[9]


for column in chain_cols:
    dam_df[column] = (
        dam_df[column]
        .fillna("")
        .astype(str)
        .str.strip()
    )


dam_df = dam_df[dam_df[family_col] != ""].copy()


# ============================================================
# RACE-RESULT SHEET
# ============================================================
# A = Place
# B = Horse
# C = Sire
# D = Dam
# E = Type of Race
# F = Distance

if race_df.shape[1] < 6:
    st.error("The Race Results sheet must contain at least 6 columns.")
    st.stop()


place_col = race_df.columns[0]
horse_col = race_df.columns[1]
sire_col = race_df.columns[2]
dam_col = race_df.columns[3]
race_type_col = race_df.columns[4]
distance_col = race_df.columns[5]


race_df = race_df[
    [
        place_col,
        horse_col,
        sire_col,
        dam_col,
        race_type_col,
        distance_col,
    ]
].copy()


race_df.columns = [
    "Place",
    "Horse",
    "Sire",
    "Dam",
    "Race_Type",
    "Distance",
]


for column in race_df.columns:
    race_df[column] = (
        race_df[column]
        .fillna("")
        .astype(str)
        .str.strip()
    )


# ============================================================
# HELPER FUNCTIONS
# ============================================================
def clean_key(value):
    """
    Creates a consistent internal key for horse names.
    """
    return re.sub(
        r"\s+",
        " ",
        str(value).strip().upper(),
    )


def classify_group_level(value):
    """
    Converts different Group-race spellings into:
    Group 1, Group 2, Group 3
    """
    text = str(value).upper()

    text = (
        text.replace("-", " ")
        .replace(".", " ")
        .replace("_", " ")
    )

    text = re.sub(r"\s+", " ", text).strip()

    if (
        "GROUP 1" in text
        or re.search(r"\bG\s*1\b", text)
        or re.search(r"\bGR\s*1\b", text)
    ):
        return "Group 1"

    if (
        "GROUP 2" in text
        or re.search(r"\bG\s*2\b", text)
        or re.search(r"\bGR\s*2\b", text)
    ):
        return "Group 2"

    if (
        "GROUP 3" in text
        or re.search(r"\bG\s*3\b", text)
        or re.search(r"\bGR\s*3\b", text)
    ):
        return "Group 3"

    return ""


def extract_place_number(value):
    """
    Extracts the numeric placing.

    Examples:
    1       -> 1
    "1st"   -> 1
    "2"     -> 2
    "3rd"   -> 3
    """
    text = str(value).strip()

    match = re.search(r"\d+", text)

    if not match:
        return None

    return int(match.group())


def distance_sort_key(value):
    """
    Sort distances numerically when possible.
    """
    match = re.search(r"\d+(?:\.\d+)?", str(value))

    if match:
        return (0, float(match.group()))

    return (1, str(value).upper())


race_df["Group_Level"] = race_df["Race_Type"].apply(
    classify_group_level
)

race_df["Place_Number"] = race_df["Place"].apply(
    extract_place_number
)


# Only rows identified as Group 1, Group 2 or Group 3
all_group_rows = race_df[
    race_df["Group_Level"].isin(
        ["Group 1", "Group 2", "Group 3"]
    )
].copy()


# ============================================================
# SIDEBAR FILTERS
# ============================================================
st.sidebar.markdown("## FILTERS")

# ------------------------------------------------------------
# TYPE OF RACE
# ------------------------------------------------------------
st.sidebar.markdown("### Type of Race")

show_group_1 = st.sidebar.checkbox(
    "Group 1",
    value=False,
)

show_group_2 = st.sidebar.checkbox(
    "Group 2",
    value=False,
)

show_group_3 = st.sidebar.checkbox(
    "Group 3",
    value=False,
)


selected_race_types = []

if show_group_1:
    selected_race_types.append("Group 1")

if show_group_2:
    selected_race_types.append("Group 2")

if show_group_3:
    selected_race_types.append("Group 3")


# ------------------------------------------------------------
# PLACING
# ------------------------------------------------------------
st.sidebar.markdown("### Placing")

placing_filter = st.sidebar.radio(
    "Select placing",
    [
        "Any placing",
        "Top 3",
        "Winner only",
    ],
    index=0,
)


# ------------------------------------------------------------
# SIRE
# ------------------------------------------------------------
available_sires = sorted(
    [
        value
        for value in all_group_rows["Sire"].unique()
        if str(value).strip()
    ],
    key=lambda value: str(value).upper(),
)


selected_sires = st.sidebar.multiselect(
    "Sire",
    available_sires,
    default=[],
    placeholder="All sires",
)


# ------------------------------------------------------------
# DISTANCE
# ------------------------------------------------------------
available_distances = sorted(
    [
        value
        for value in all_group_rows["Distance"].unique()
        if str(value).strip()
    ],
    key=distance_sort_key,
)


selected_distances = st.sidebar.multiselect(
    "Distance",
    available_distances,
    default=[],
    placeholder="All distances",
)


# ------------------------------------------------------------
# DISPLAY CONTROLS
# These control the drawing, not the data filtering.
# ------------------------------------------------------------
st.sidebar.markdown("---")
st.sidebar.markdown("## DISPLAY")

show_names = st.sidebar.checkbox(
    "Show Names",
    value=True,
)

show_rings = st.sidebar.checkbox(
    "Show Generation Rings",
    value=True,
)

ring_spacing = st.sidebar.slider(
    "Generation Ring Spacing",
    120,
    360,
    200,
    5,
)

label_gap_from_node = st.sidebar.slider(
    "Name Gap Before Node",
    5,
    100,
    18,
    1,
)

font_size = st.sidebar.slider(
    "Label Font Size",
    8,
    18,
    12,
    1,
)

canvas_height = st.sidebar.slider(
    "Canvas Height",
    700,
    1800,
    1100,
    50,
)

rotation_degrees = st.sidebar.slider(
    "Rotate Tree",
    -180,
    180,
    0,
    1,
)


st.sidebar.markdown("---")
st.sidebar.markdown("## LEGEND")
st.sidebar.markdown("🔴 Foundation mare")
st.sidebar.markdown("🟠 Connecting maternal ancestor")
st.sidebar.markdown("⚫ Matching runner")
st.sidebar.markdown("🔵 Dam of a matching runner")


# ============================================================
# DETERMINE WHETHER PERFORMANCE FILTERS ARE ACTIVE
# ============================================================
performance_filters_active = any(
    [
        bool(selected_race_types),
        placing_filter != "Any placing",
        bool(selected_sires),
        bool(selected_distances),
    ]
)


# ============================================================
# FILTER RACE RESULTS
# ============================================================
filtered_race_df = all_group_rows.copy()


# If no Group checkbox is selected but another filter is active,
# use all three Group levels.
if selected_race_types:
    filtered_race_df = filtered_race_df[
        filtered_race_df["Group_Level"].isin(
            selected_race_types
        )
    ]


if placing_filter == "Winner only":
    filtered_race_df = filtered_race_df[
        filtered_race_df["Place_Number"] == 1
    ]

elif placing_filter == "Top 3":
    filtered_race_df = filtered_race_df[
        filtered_race_df["Place_Number"].isin(
            [1, 2, 3]
        )
    ]


if selected_sires:
    filtered_race_df = filtered_race_df[
        filtered_race_df["Sire"].isin(
            selected_sires
        )
    ]


if selected_distances:
    filtered_race_df = filtered_race_df[
        filtered_race_df["Distance"].isin(
            selected_distances
        )
    ]


# ============================================================
# DYNAMIC FAMILY DROPDOWN
# ============================================================
# Total population of each family in the Dam Lines sheet.
family_counts = (
    dam_df[family_col]
    .dropna()
    .astype(str)
    .str.strip()
    .value_counts()
)


# Build a lookup from every name in the dam-line chains
# to the family or families in which that name appears.
name_to_families = defaultdict(set)

for _, row in dam_df.iterrows():
    family_name = str(row[family_col]).strip()

    if not family_name:
        continue

    for column in chain_cols:
        name_key = clean_key(row[column])

        if name_key:
            name_to_families[name_key].add(
                family_name
            )


# When performance filters are active, count unique matching
# runners in each family. A race result belongs to a family
# when either the runner or its dam appears in that family tree.
matching_runners_by_family = defaultdict(set)

if performance_filters_active:

    for _, row in filtered_race_df.iterrows():
        horse_key = clean_key(row["Horse"])
        dam_key = clean_key(row["Dam"])

        matching_families = (
            name_to_families.get(
                horse_key,
                set(),
            )
            | name_to_families.get(
                dam_key,
                set(),
            )
        )

        for family_name in matching_families:
            if horse_key:
                matching_runners_by_family[
                    family_name
                ].add(horse_key)


    # Keep only families with at least one matching runner.
    # Sort by matching-runner count, then total population,
    # then alphabetically.
    families = sorted(
        matching_runners_by_family.keys(),
        key=lambda family_name: (
            -len(
                matching_runners_by_family[
                    family_name
                ]
            ),
            -int(
                family_counts.get(
                    family_name,
                    0,
                )
            ),
            family_name.upper(),
        ),
    )

else:

    # No performance filters:
    # show all families, largest population first.
    families = family_counts.index.tolist()


if not families:
    st.sidebar.warning(
        "No families match the selected filters."
    )
    st.warning(
        "No families contain runners matching the selected "
        "Group, placing, sire and distance filters."
    )
    st.stop()


def family_dropdown_label(family_name):
    # Show matching-runner count while filters are active.
    # Otherwise show total family population.
    if performance_filters_active:
        matching_count = len(
            matching_runners_by_family[
                family_name
            ]
        )

        return (
            f"{family_name} "
            f"({matching_count} matching)"
        )

    return (
        f"{family_name} "
        f"({int(family_counts[family_name])})"
    )


selected_family = st.sidebar.selectbox(
    "Family",
    families,
    format_func=family_dropdown_label,
)


# ============================================================
# SELECTED FAMILY
# ============================================================
family_df = dam_df[
    dam_df[family_col] == selected_family
].copy()


# ============================================================
# BUILD THE BASE DAM-LINE TREE
# ============================================================
children = defaultdict(set)
parent_of = {}
display_names = {}
nodes_set = set()


for _, row in family_df.iterrows():

    raw_chain = [
        row[column]
        for column in chain_cols
        if str(row[column]).strip()
    ]

    if not raw_chain:
        continue

    # Spreadsheet runs from current mare backwards.
    # Reverse it so the tree runs:
    # foundation mare -> younger generation
    raw_chain = list(reversed(raw_chain))

    chain = []

    for name in raw_chain:
        key = clean_key(name)

        if not key:
            continue

        chain.append(key)
        nodes_set.add(key)

        display_names.setdefault(
            key,
            str(name).strip(),
        )

    for index in range(len(chain) - 1):
        parent = chain[index]
        child = chain[index + 1]

        if parent == child:
            continue

        children[parent].add(child)

        # The dam-line data should normally give one mother per mare.
        parent_of.setdefault(child, parent)


root = clean_key(selected_family)

display_names[root] = selected_family
nodes_set.add(root)


# Nodes belonging to the original family tree,
# before adding race-result progeny
family_node_keys = set(nodes_set)


# ============================================================
# FIND MATCHING RUNNERS AND PRODUCERS IN THIS FAMILY
# ============================================================
matching_runner_keys = set()
matching_producer_keys = set()

relevant_race_rows = []


if performance_filters_active:

    for row_index, row in filtered_race_df.iterrows():

        horse_name = row["Horse"]
        dam_name = row["Dam"]

        horse_key = clean_key(horse_name)
        dam_key = clean_key(dam_name)

        if not horse_key:
            continue

        horse_is_in_family_tree = (
            horse_key in family_node_keys
        )

        dam_is_in_family_tree = (
            dam_key in family_node_keys
        )

        # A result is relevant to the selected family when:
        # 1. the runner is already a mare/horse in the family tree, or
        # 2. its dam belongs to the selected family.
        if not horse_is_in_family_tree and not dam_is_in_family_tree:
            continue

        matching_runner_keys.add(horse_key)
        relevant_race_rows.append(row_index)

        display_names.setdefault(
            horse_key,
            horse_name,
        )

        # Add a runner not already present in the dam-line tree
        # beneath its dam.
        if dam_is_in_family_tree:
            children[dam_key].add(horse_key)
            parent_of.setdefault(horse_key, dam_key)
            nodes_set.add(horse_key)
            matching_producer_keys.add(dam_key)


# Remove duplicate race-result rows from the displayed table
if relevant_race_rows:
    relevant_results_df = (
        filtered_race_df
        .loc[relevant_race_rows]
        .copy()
    )
else:
    relevant_results_df = filtered_race_df.iloc[0:0].copy()


children = {
    key: sorted(values)
    for key, values in children.items()
}


# ============================================================
# DETERMINE VISIBLE NODES
# ============================================================
if not performance_filters_active:

    # No performance filter:
    # show the complete family tree.
    visible_nodes = set(nodes_set)

else:

    # Filters active:
    # show only matching runners,
    # their dams/producers,
    # and every maternal ancestor required to connect them
    # to the foundation mare.
    visible_nodes = {root}

    matched_nodes = (
        matching_runner_keys
        | matching_producer_keys
    )

    matched_nodes = {
        node_key
        for node_key in matched_nodes
        if node_key in nodes_set
    }

    for matched_node in matched_nodes:

        current_node = matched_node
        visited_in_chain = set()

        while current_node:

            if current_node in visited_in_chain:
                break

            visited_in_chain.add(current_node)
            visible_nodes.add(current_node)

            if current_node == root:
                break

            current_node = parent_of.get(current_node)


# ============================================================
# BUILD VISIBLE PARENT-CHILD LINKS
# ============================================================
visible_children = defaultdict(list)


for parent, child_set in children.items():

    if parent not in visible_nodes:
        continue

    for child in child_set:

        if child in visible_nodes:
            visible_children[parent].append(child)


visible_children = {
    parent: sorted(set(child_list))
    for parent, child_list in visible_children.items()
}


# ============================================================
# RADIAL POSITIONING
# ============================================================
positions = {
    root: {
        "x": 0,
        "y": 0,
        "depth": 0,
        "angle": 0,
    }
}


def count_leaves(node_key, visiting=None):

    if visiting is None:
        visiting = set()

    if node_key in visiting:
        return 1

    visiting = set(visiting)
    visiting.add(node_key)

    child_list = visible_children.get(
        node_key,
        [],
    )

    if not child_list:
        return 1

    return sum(
        count_leaves(child, visiting)
        for child in child_list
    )


def subtree_size(node_key, visiting=None):

    if visiting is None:
        visiting = set()

    if node_key in visiting:
        return 1

    visiting = set(visiting)
    visiting.add(node_key)

    child_list = visible_children.get(
        node_key,
        [],
    )

    return 1 + sum(
        subtree_size(child, visiting)
        for child in child_list
    )


def assign_positions(
    node_key,
    start_angle,
    end_angle,
    depth,
    visiting=None,
):

    if visiting is None:
        visiting = set()

    if node_key in visiting:
        return

    visiting = set(visiting)
    visiting.add(node_key)

    child_list = visible_children.get(
        node_key,
        [],
    )

    if not child_list:
        return

    weights = [
        max(
            count_leaves(child),
            math.sqrt(subtree_size(child)),
        )
        for child in child_list
    ]

    total_weight = sum(weights)

    if total_weight <= 0:
        return

    current_angle = start_angle

    for child, weight in zip(
        child_list,
        weights,
    ):

        portion = weight / total_weight

        child_start = current_angle

        child_end = (
            current_angle
            + (end_angle - start_angle) * portion
        )

        angle = (
            child_start + child_end
        ) / 2

        radius = depth * ring_spacing

        positions[child] = {
            "x": radius * math.cos(angle),
            "y": radius * math.sin(angle),
            "depth": depth,
            "angle": angle,
        }

        assign_positions(
            child,
            child_start,
            child_end,
            depth + 1,
            visiting,
        )

        current_angle = child_end


assign_positions(
    root,
    -math.pi / 2,
    1.5 * math.pi,
    1,
)


# Normally all connected visible nodes receive positions.
# This fallback protects against incomplete source data.
for node_key in visible_nodes:

    if node_key not in positions:
        positions[node_key] = {
            "x": 0,
            "y": 0,
            "depth": 1,
            "angle": 0,
        }


max_depth = max(
    [
        position["depth"]
        for position in positions.values()
    ]
    or [1]
)


# ============================================================
# PREPARE NODES AND EDGES
# ============================================================
nodes = []


for node_key in visible_nodes:

    position = positions[node_key]

    is_matching_runner = (
        node_key in matching_runner_keys
    )

    is_matching_producer = (
        node_key in matching_producer_keys
    )

    if is_matching_runner:
        color = "#000000"

    elif is_matching_producer:
        color = "#1F77B4"

    elif node_key == root:
        color = "#7B1E3A"

    else:
        color = "#E49A12"

    nodes.append(
        {
            "id": node_key,
            "label": display_names.get(
                node_key,
                node_key,
            ),
            "x": position["x"],
            "y": position["y"],
            "angle": position["angle"],
            "depth": position["depth"],
            "root": node_key == root,
            "color": color,
        }
    )


edges = []


for parent, child_list in visible_children.items():

    for child in child_list:

        if (
            parent in positions
            and child in positions
        ):
            edges.append(
                {
                    "source": parent,
                    "target": child,
                }
            )


tree_data = {
    "nodes": nodes,
    "edges": edges,
    "root": root,
    "maxDepth": max_depth,
    "showNames": show_names,
    "showRings": show_rings,
    "ringSpacing": ring_spacing,
    "labelGapFromNode": label_gap_from_node,
    "fontSize": font_size,
    "canvasHeight": canvas_height,
    "rotationDegrees": rotation_degrees,
}


# ============================================================
# RUNNER COUNT
# ============================================================
matching_runner_count = len(
    {
        clean_key(value)
        for value in relevant_results_df["Horse"]
        if clean_key(value)
    }
)


# ============================================================
# PAGE SUMMARY
# ============================================================
summary_col_1, summary_col_2 = st.columns(2)

with summary_col_1:
    st.metric(
        "Matching runners",
        matching_runner_count
        if performance_filters_active
        else "All",
    )

with summary_col_2:
    if performance_filters_active:
        active_group_text = (
            ", ".join(selected_race_types)
            if selected_race_types
            else "Group 1, Group 2 and Group 3"
        )

        st.metric(
            "Race levels searched",
            active_group_text,
        )
    else:
        st.metric(
            "Tree view",
            "Full family",
        )


# ============================================================
# HTML / JAVASCRIPT TREE
# ============================================================
html = """
<!DOCTYPE html>
<html>
<head>
<script src="https://cdn.jsdelivr.net/npm/d3@7"></script>

<style>
body {
  margin: 0;
  background: #fbf7ef;
  font-family: Arial, sans-serif;
}

#container {
  width: 100%;
  height: __CANVAS_HEIGHT__px;
  background:
    radial-gradient(
      circle,
      #fffaf1 0%,
      #fbf7ef 85%
    );
  border: 1px solid #e6d8c3;
  overflow: hidden;
}

.link {
  stroke: #c9a46a;
  stroke-width: 1.1px;
  fill: none;
  opacity: 0.85;
}

.ring {
  fill: none;
  stroke: #ddc59e;
  stroke-width: 1px;
  stroke-dasharray: 4 4;
  opacity: 0.7;
}

.node circle {
  stroke: #8a6b3f;
  stroke-width: 1.4px;
}

.label {
  fill: #111;
  paint-order: stroke;
  stroke: #fbf7ef;
  stroke-width: 5px;
  stroke-linejoin: round;
  font-weight: 600;
}

.root-label {
  font-size: 20px;
  font-weight: bold;
}
</style>
</head>

<body>
<div id="container"></div>

<script>
const data = __TREE_DATA__;

const container = document.getElementById("container");
const width = container.clientWidth;
const height = data.canvasHeight;

const radiusStep = data.ringSpacing;
const fontSize = data.fontSize;
const labelGap = data.labelGapFromNode;
const rotationDegrees = data.rotationDegrees;

const svg = d3
  .select("#container")
  .append("svg")
  .attr("width", width)
  .attr("height", height);

const zoomLayer = svg
  .append("g")
  .attr(
    "transform",
    `translate(${width / 2},${height / 2})`
  );

const rotationLayer = zoomLayer
  .append("g")
  .attr(
    "transform",
    `rotate(${rotationDegrees})`
  );

svg.call(
  d3
    .zoom()
    .scaleExtent([0.15, 5])
    .on("zoom", function(event) {
      zoomLayer.attr(
        "transform",
        event.transform
      );
    })
);

const nodeById = {};

data.nodes.forEach(function(node) {
  nodeById[node.id] = node;
});

const parentByChild = {};

data.edges.forEach(function(edge) {
  parentByChild[edge.target] = edge.source;
});

function nodeRadius(node) {
  if (node.root) {
    return 24;
  }

  return Math.max(
    5,
    15 - node.depth
  );
}


// ------------------------------------------------------------
// GENERATION RINGS
// ------------------------------------------------------------
if (data.showRings) {

  for (
    let generation = 1;
    generation <= data.maxDepth;
    generation++
  ) {

    rotationLayer
      .append("circle")
      .attr("class", "ring")
      .attr(
        "r",
        generation * radiusStep
      );

    rotationLayer
      .append("text")
      .attr("x", 8)
      .attr(
        "y",
        -generation * radiusStep + 14
      )
      .attr("font-size", "11px")
      .attr("fill", "#7B1E3A")
      .text("GEN " + generation);
  }
}


// ------------------------------------------------------------
// CONNECTION LINES
// ------------------------------------------------------------
data.edges.forEach(function(edge) {

  const source = nodeById[edge.source];
  const target = nodeById[edge.target];

  if (!source || !target) {
    return;
  }

  rotationLayer
    .append("line")
    .attr("class", "link")
    .attr("x1", source.x)
    .attr("y1", source.y)
    .attr("x2", target.x)
    .attr("y2", target.y);
});


// ------------------------------------------------------------
// NODES
// ------------------------------------------------------------
const nodeGroups = rotationLayer
  .selectAll(".node")
  .data(data.nodes)
  .enter()
  .append("g")
  .attr("class", "node")
  .attr(
    "transform",
    function(node) {
      return `translate(${node.x},${node.y})`;
    }
  );

nodeGroups
  .append("circle")
  .attr(
    "r",
    function(node) {
      return nodeRadius(node);
    }
  )
  .attr(
    "fill",
    function(node) {
      return node.color;
    }
  );


// ------------------------------------------------------------
// LABELS
// ------------------------------------------------------------
if (data.showNames) {

  rotationLayer
    .selectAll(".label")
    .data(data.nodes)
    .enter()
    .append("text")
    .attr(
      "class",
      function(node) {
        return node.root
          ? "label root-label"
          : "label";
      }
    )
    .attr(
      "font-size",
      function(node) {
        return node.root
          ? "20px"
          : fontSize + "px";
      }
    )
    .attr(
      "dominant-baseline",
      "middle"
    )
    .attr(
      "text-anchor",
      function(node) {
        return node.root
          ? "middle"
          : "end";
      }
    )
    .attr(
      "transform",
      function(node) {

        if (node.root) {
          return (
            `translate(${node.x},${node.y + 44})`
          );
        }

        const parentId =
          parentByChild[node.id];

        const parent =
          nodeById[parentId];

        if (!parent) {
          return (
            `translate(${node.x},${node.y})`
          );
        }

        const dx =
          node.x - parent.x;

        const dy =
          node.y - parent.y;

        const angle =
          Math.atan2(dy, dx);

        const radius =
          nodeRadius(node);

        const endX =
          node.x
          - Math.cos(angle)
          * (radius + labelGap);

        const endY =
          node.y
          - Math.sin(angle)
          * (radius + labelGap);

        // Names are intentionally not flipped.
        // The entire tree can be rotated by the user.
        const rotation =
          angle * 180 / Math.PI;

        return (
          `translate(${endX},${endY}) `
          + `rotate(${rotation})`
        );
      }
    )
    .text(
      function(node) {
        return node.label;
      }
    );
}
</script>
</body>
</html>
"""


html = html.replace(
    "__CANVAS_HEIGHT__",
    str(canvas_height),
)

html = html.replace(
    "__TREE_DATA__",
    json.dumps(tree_data),
)


# ============================================================
# DISPLAY TREE
# ============================================================
st.subheader(f"Family: {selected_family}")

components.html(
    html,
    height=canvas_height + 20,
)


# ============================================================
# FILTERED RESULTS TABLE
# ============================================================
st.markdown("### Matching race results")

if performance_filters_active:

    if relevant_results_df.empty:
        st.warning(
            "No runners in this family match the selected filters."
        )

    else:
        display_result_columns = [
            "Place",
            "Horse",
            "Sire",
            "Dam",
            "Race_Type",
            "Distance",
            "Group_Level",
        ]

        st.dataframe(
            relevant_results_df[
                display_result_columns
            ],
            use_container_width=True,
            hide_index=True,
        )

else:
    st.info(
        "No performance filters are active. "
        "The complete family tree is displayed."
    )


# ============================================================
# DAM-LINE DATA TABLE
# ============================================================
st.markdown(
    f"### Dam-line data for “{selected_family}”"
)

st.caption(
    f"Total rows in this family: {len(family_df)}"
)

st.dataframe(
    family_df[chain_cols],
    use_container_width=True,
    hide_index=True,
)
