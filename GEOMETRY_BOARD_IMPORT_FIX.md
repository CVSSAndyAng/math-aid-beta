# Geometry board NameError fix

The deployed app raised:

`NameError: name 'st_components_v1' is not defined`

The geometry construction board is registered with Streamlit Components V1. This build now imports
`streamlit.components.v1` both at module level and immediately beside the custom-component declaration,
so the component registration cannot reference an undefined alias after code merges/reordering.

The `geometry_board_component/` frontend remains bundled with the app.
