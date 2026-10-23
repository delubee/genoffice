import { valueAxisScale, type ChartVisualState } from '@genoffice/xlsx-gateway/domain/chart-visual'

// These are authored chart colors, independent of editor chrome themes.
const palette = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47']

/** The four MVP chart types share the gateway's series model and axis scale. */
export function BasicChart({ chart }: { chart: ChartVisualState }) {
  const series = chart.series,
    values = series.flatMap((item) => item.values),
    scale = valueAxisScale(Math.max(0, ...values), chart.valueAxis, Math.min(0, ...values)),
    count = Math.max(1, ...series.map((item) => item.values.length)),
    width = 360,
    height = 200,
    x = (value: number) => 55 + ((value - scale.min) / (scale.max - scale.min)) * width,
    y = (value: number) => 250 - ((value - scale.min) / (scale.max - scale.min)) * height,
    color = (index: number) => series[index]?.color ?? palette[index % palette.length],
    horizontal = chart.barDirection === 'bar',
    pie = chart.chartTypes.includes('pieChart'),
    line = chart.chartTypes.includes('lineChart')
  let angle = -Math.PI / 2
  const total = series[0]?.values.reduce((sum, value) => sum + Math.abs(value), 0) ?? 0
  return (
    <svg viewBox="0 0 480 320" role="img" aria-label={chart.title}>
      <title>{chart.title}</title>
      <rect width="480" height="320" fill="#FFFFFF" />
      <text x="240" y="25" textAnchor="middle" fill="#000000" fontSize="16">
        {chart.title}
      </text>
      {pie ? (
        series[0]?.values.map((value, index) => {
          if (!total || !value) return null
          const start = angle,
            delta = (Math.abs(value) / total) * Math.PI * 2
          angle += delta
          const fill =
            series[0]?.pointColors?.find((point) => point.index === index)?.color ??
            palette[index % palette.length]
          return delta >= Math.PI * 2 - 1e-9 ? (
            <circle key={index} cx="240" cy="150" r="95" fill={fill}>
              <title>
                {series[0]?.categories[index]}: {value}
              </title>
            </circle>
          ) : (
            <path
              key={index}
              d={`M240,150 L${240 + Math.cos(start) * 95},${150 + Math.sin(start) * 95} A95,95 0 ${delta > Math.PI ? 1 : 0},1 ${240 + Math.cos(angle) * 95},${150 + Math.sin(angle) * 95} Z`}
              fill={fill}
            >
              <title>
                {series[0]?.categories[index]}: {value}
              </title>
            </path>
          )
        })
      ) : (
        <>
          {scale.ticks.map((tick) =>
            horizontal ? (
              <g key={tick}>
                <line x1={x(tick)} x2={x(tick)} y1="50" y2="250" stroke="#DDDDDD" />
                <text x={x(tick)} y="267" textAnchor="middle" fill="#000000" fontSize="10">
                  {tick}
                </text>
              </g>
            ) : (
              <g key={tick}>
                <line x1="55" x2="415" y1={y(tick)} y2={y(tick)} stroke="#DDDDDD" />
                <text x="50" y={y(tick) + 3} textAnchor="end" fill="#000000" fontSize="10">
                  {tick}
                </text>
              </g>
            ),
          )}
          {series.map((item, index) =>
            line ? (
              <g key={index}>
                {item.values.map((value, point) =>
                  item.blanks?.includes(point) ? null : (
                    <g key={point}>
                      {point > 0 && !item.blanks?.includes(point - 1) && (
                        <line
                          x1={55 + ((point - 0.5) / count) * width}
                          y1={y(item.values[point - 1]!)}
                          x2={55 + ((point + 0.5) / count) * width}
                          y2={y(value)}
                          stroke={color(index)}
                          strokeWidth="2"
                        />
                      )}
                      <circle
                        cx={55 + ((point + 0.5) / count) * width}
                        cy={y(value)}
                        r="3"
                        fill={color(index)}
                      >
                        <title>
                          {item.name}, {item.categories[point]}: {value}
                        </title>
                      </circle>
                    </g>
                  ),
                )}
              </g>
            ) : (
              <g key={index}>
                {item.values.map((value, point) => {
                  if (item.blanks?.includes(point)) return null
                  const band = (horizontal ? height : width) / count,
                    thickness = (band * 0.75) / series.length,
                    start = (horizontal ? 50 : 55) + band * point + band * 0.125 + index * thickness
                  return horizontal ? (
                    <rect
                      key={point}
                      x={Math.min(x(0), x(value))}
                      y={start}
                      width={Math.abs(x(value) - x(0))}
                      height={thickness}
                      fill={color(index)}
                    >
                      <title>
                        {item.name}, {item.categories[point]}: {value}
                      </title>
                    </rect>
                  ) : (
                    <rect
                      key={point}
                      x={start}
                      y={Math.min(y(0), y(value))}
                      width={thickness}
                      height={Math.abs(y(value) - y(0))}
                      fill={color(index)}
                    >
                      <title>
                        {item.name}, {item.categories[point]}: {value}
                      </title>
                    </rect>
                  )
                })}
              </g>
            ),
          )}
          {Array.from({ length: count }, (_, point) => point)
            .filter((point) => point % Math.max(1, Math.ceil(count / 8)) === 0)
            .map((point) =>
              horizontal ? (
                <text
                  key={point}
                  x="50"
                  y={50 + ((point + 0.5) / count) * height}
                  textAnchor="end"
                  fill="#000000"
                  fontSize="10"
                >
                  {series[0]?.categories[point]?.slice(0, 8)}
                </text>
              ) : (
                <text
                  key={point}
                  x={55 + ((point + 0.5) / count) * width}
                  y="267"
                  textAnchor="middle"
                  fill="#000000"
                  fontSize="10"
                >
                  {series[0]?.categories[point]?.slice(0, 8)}
                </text>
              ),
            )}
        </>
      )}
      {chart.legend !== 'none' &&
        (pie
          ? (series[0]?.categories ?? []).slice(0, 6)
          : series.map((item) => item.name).slice(0, 6)
        ).map((name, index) => (
          <g key={index}>
            <rect
              x={15 + (index % 3) * 155}
              y={280 + Math.floor(index / 3) * 18}
              width="9"
              height="9"
              fill={pie ? palette[index % palette.length] : color(index)}
            />
            <text
              x={28 + (index % 3) * 155}
              y={288 + Math.floor(index / 3) * 18}
              fill="#000000"
              fontSize="10"
            >
              {name.slice(0, 20)}
            </text>
          </g>
        ))}
    </svg>
  )
}
