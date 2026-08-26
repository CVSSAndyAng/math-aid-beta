# Context-image relevance filter

Paper Setter now suppresses generic context images when they contain or are likely to
contain more objects than the mathematical question describes.

In particular:
- two similar jugs/containers -> exactly two clean similar-object drawings;
- one mathematical object -> no stock image containing several unrelated objects;
- geometry/mensuration -> deterministic mathematical diagram is preferred;
- if no sufficiently relevant image is available, no context image is shown.

This prevents images such as a box + jar + can + pitcher from being used for a question
about two similar cylindrical containers.
