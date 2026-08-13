// 兼容层：原 person 工具已上移到 src/shared/person，供主/渲染进程共用。
// 这里统一转导，避免改动已有引用方（ResultList 等）。
export * from '../../shared/person'
