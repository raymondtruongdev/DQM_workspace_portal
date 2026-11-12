import * as React from 'react';
import {
  AddRegular,
  AppsRegular,
  ArrowDownloadRegular,
  ArrowLeftRegular,
  ArrowRightRegular,
  ArrowSyncRegular,
  ArrowUploadRegular,
  BuildingRegular,
  CalendarDayRegular,
  CalendarRegular,
  CheckmarkCircleRegular,
  CheckmarkRegular,
  ChevronDownRegular,
  ChevronRightRegular,
  ChevronUpRegular,
  CopyRegular,
  DismissRegular,
  DocumentArrowUpRegular,
  ErrorCircleRegular,
  EyeRegular,
  GridRegular,
  HomeRegular,
  ImageRegular,
  InfoRegular,
  LockClosedRegular,
  MailRegular,
  PaintBrushRegular,
  PersonRegular,
  PowerRegular,
  PrintRegular,
  SaveRegular,
  SearchRegular,
  TableRegular,
  WarningRegular,
} from '@fluentui/react-icons';

/**
 * Bộ icon dùng chung của toàn app – Fluent UI System Icons (Regular).
 * Mọi icon đi qua component này để đồng nhất nét vẽ và kích thước.
 * Kích thước chuẩn: 16 (trong nút nhỏ / chữ), 20 (nút thường), 24 (ô thẻ).
 */
const ICONS = {
  add: AddRegular,
  apps: AppsRegular,
  'arrow-left': ArrowLeftRegular,
  'arrow-right': ArrowRightRegular,
  building: BuildingRegular,
  calendar: CalendarRegular,
  'calendar-day': CalendarDayRegular,
  check: CheckmarkRegular,
  'check-circle': CheckmarkCircleRegular,
  'chevron-down': ChevronDownRegular,
  'chevron-right': ChevronRightRegular,
  'chevron-up': ChevronUpRegular,
  copy: CopyRegular,
  dismiss: DismissRegular,
  download: ArrowDownloadRegular,
  'error-circle': ErrorCircleRegular,
  eye: EyeRegular,
  'file-up': DocumentArrowUpRegular,
  grid: GridRegular,
  home: HomeRegular,
  image: ImageRegular,
  info: InfoRegular,
  lock: LockClosedRegular,
  mail: MailRegular,
  paint: PaintBrushRegular,
  person: PersonRegular,
  power: PowerRegular,
  print: PrintRegular,
  refresh: ArrowSyncRegular,
  save: SaveRegular,
  search: SearchRegular,
  table: TableRegular,
  upload: ArrowUploadRegular,
  warning: WarningRegular,
} as const;

export type IconName = keyof typeof ICONS;

/** Danh sách tên icon hiện có (dùng cho trang UI Kit). */
export const ICON_NAMES = Object.keys(ICONS) as IconName[];

export interface IconProps {
  name: IconName;
  /** Kích thước px (mặc định 20). */
  size?: number;
  className?: string;
  /** Xoay liên tục – dùng cho trạng thái đang tải. */
  spin?: boolean;
}

/**
 * Icon dùng chung (Fluent UI Regular). Không viết <svg> tay.
 * @example <Icon name="save" />  ·  <Icon name="refresh" size={16} spin={isLoading} />
 */
export const Icon: React.FC<IconProps> = ({ name, size = 20, className, spin }) => {
  const Glyph = ICONS[name];
  const cls = ['icon', spin ? 'spin' : '', className || ''].filter(Boolean).join(' ');
  return <Glyph className={cls} style={{ width: size, height: size }} aria-hidden="true" focusable="false" />;
};
