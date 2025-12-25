import { ImageResponse } from "next/og";

export const runtime = "edge";

export const alt = "OpenSheets - The open source spreadsheet agent";
export const size = {
  width: 1200,
  height: 630,
};
export const contentType = "image/png";

async function loadGoogleFont(font: string) {
  const url = `https://fonts.googleapis.com/css2?family=${font.replace(/ /g, "+")}:wght@400`;
  const css = await (
    await fetch(url, {
      headers: {
        "User-Agent":
          "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_6_8; de-at) AppleWebKit/533.21.1 (KHTML, like Gecko) Version/5.0.5 Safari/533.21.1",
      },
    })
  ).text();

  const resource = css.match(
    /src: url\((.+)\) format\('(opentype|truetype)'\)/,
  );

  if (resource) {
    const response = await fetch(resource[1]);
    if (response.status === 200) {
      return await response.arrayBuffer();
    }
  }

  throw new Error("Failed to load font");
}

export default async function Image() {
  const instrumentSerif = await loadGoogleFont("Instrument Serif");

  return new ImageResponse(
    <div
      style={{
        width: "100%",
        height: "100%",
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        justifyContent: "center",
        backgroundColor: "#fff",
        position: "relative",
      }}
    >
      {/* Grid pattern background */}
      <div
        style={{
          position: "absolute",
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          display: "flex",
          flexDirection: "column",
        }}
      >
        {/* Horizontal lines */}
        {Array.from({ length: 16 }).map((_, i) => (
          <div
            key={`h-${i}`}
            style={{
              position: "absolute",
              top: i * 42,
              left: 0,
              right: 0,
              height: 1,
              backgroundColor: i === 0 ? "#c5cbb8" : "#e8ebe2",
            }}
          />
        ))}
        {/* Vertical lines */}
        {Array.from({ length: 13 }).map((_, i) => (
          <div
            key={`v-${i}`}
            style={{
              position: "absolute",
              top: 0,
              bottom: 0,
              left: i === 0 ? 60 : 60 + i * 95,
              width: 1,
              backgroundColor: i === 0 ? "#c5cbb8" : "#e8ebe2",
            }}
          />
        ))}
      </div>

      {/* Row numbers column */}
      <div
        style={{
          position: "absolute",
          top: 42,
          left: 0,
          width: 60,
          bottom: 0,
          display: "flex",
          flexDirection: "column",
        }}
      >
        {Array.from({ length: 14 }).map((_, i) => (
          <div
            key={`row-${i}`}
            style={{
              height: 42,
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              fontSize: 13,
              color: "#8b9178",
              fontFamily: "system-ui",
            }}
          >
            {i + 1}
          </div>
        ))}
      </div>

      {/* Column headers */}
      <div
        style={{
          position: "absolute",
          top: 0,
          left: 60,
          right: 0,
          height: 42,
          display: "flex",
        }}
      >
        {["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"].map(
          (letter) => (
            <div
              key={letter}
              style={{
                width: 95,
                height: 42,
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                fontSize: 13,
                color: "#8b9178",
                fontFamily: "system-ui",
              }}
            >
              {letter}
            </div>
          ),
        )}
      </div>

      {/* Content overlay with gradient */}
      <div
        style={{
          position: "absolute",
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          background:
            "radial-gradient(ellipse 60% 50% at 50% 50%, rgba(255,255,255,0.92) 0%, rgba(255,255,255,0.5) 60%, rgba(255,255,255,0) 100%)",
          display: "flex",
          flexDirection: "column",
          alignItems: "center",
          justifyContent: "center",
          gap: 24,
        }}
      >
        {/* Title */}
        <div
          style={{
            fontSize: 96,
            fontFamily: "Instrument Serif",
            color: "#1a1d16",
            letterSpacing: "-0.02em",
          }}
        >
          OpenSheets
        </div>

        {/* Description */}
        <div
          style={{
            fontSize: 32,
            fontFamily: "Instrument Serif",
            color: "#596352",
            letterSpacing: "0.01em",
          }}
        >
          The open source spreadsheet agent
        </div>
      </div>
    </div>,
    {
      ...size,
      fonts: [
        {
          name: "Instrument Serif",
          data: instrumentSerif,
          style: "normal",
          weight: 400,
        },
      ],
    },
  );
}
